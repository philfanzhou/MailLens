# Install from a fixed Windows checkout; status/uninstall need only this script.
[CmdletBinding()]
param(
    [ValidateSet('Install', 'Status', 'Update', 'Uninstall')][string]$Action = 'Install',
    [switch]$User,
    [string]$Project,
    [string]$SourcePath = (Split-Path $PSScriptRoot -Parent)
)
$ErrorActionPreference = 'Stop'
$script:InstallerDefaultSource = Split-Path $PSScriptRoot -Parent
$script:InstallPhaseHook = $null
# Windows PowerShell 5.1 / built-in .NET installer. No source files are executed.
Set-StrictMode -Version Latest
$script:ReceiptName = '.mail-lens-install.json'

function Get-InstallAbsolutePath {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path) -or $Path -notmatch '^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+(?:[\\/]|$))') {
        throw "$Label must be an absolute Windows path."
    }
    if ($Path -match '[\x00-\x1f<>"|?*]' -or $Path.Substring(2).Contains(':')) { throw "$Label contains invalid path characters." }
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -eq [IO.Path]::GetPathRoot($full)) { return $full }
    return $full.TrimEnd('\', '/')
}

function Test-InstallContained {
    param([string]$Parent, [string]$Child)
    return $Child.Equals($Parent, [StringComparison]::OrdinalIgnoreCase) -or $Child.StartsWith($Parent.TrimEnd('\', '/') + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Assert-InstallSafePath {
    param([string]$Path)
    $current = $Path
    while (-not [string]::IsNullOrEmpty($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Links or reparse points are not supported: $current" }
            if ($current -ne $Path -and -not $item.PSIsContainer) { throw "A path ancestor is not a directory: $current" }
        }
        $parent = [IO.Path]::GetDirectoryName($current)
        if ($parent -eq $current) { break }
        $current = $parent
    }
}

function Assert-InstallRelativePath {
    param($Path)
    if ($Path -isnot [string] -or [string]::IsNullOrEmpty($Path) -or $Path.Contains('\')) { throw 'Invalid manifest path.' }
    foreach ($part in $Path.Split('/')) {
        if ([string]::IsNullOrEmpty($part) -or $part -in @('.', '..') -or $part -match '[\x00-\x1f<>:"\\|?*]' -or $part -match '[. ]$' -or $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { throw 'Unsafe manifest path.' }
    }
}

function Get-InstallHash {
    param([string]$Path)
    $stream = [IO.File]::OpenRead($Path)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-', '').ToLowerInvariant() }
    finally { $hash.Dispose(); $stream.Dispose() }
}

function Get-InstallInventory {
    param([string]$Root, [switch]$Installed)
    Assert-InstallSafePath $Root
    if (-not [IO.Directory]::Exists($Root)) { throw "Package directory is missing: $Root" }
    $files = New-Object 'System.Collections.Generic.List[object]'
    $directories = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $pending = New-Object 'System.Collections.Generic.Queue[string]'
    $pending.Enqueue($Root)
    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        foreach ($item in @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Package contains a link or reparse point: $($item.FullName)" }
            $relative = $item.FullName.Substring($Root.Length + 1).Replace('\', '/')
            Assert-InstallRelativePath $relative
            if (-not $seen.Add($relative)) { throw 'Case-colliding package paths.' }
            if ($relative -eq $script:ReceiptName) {
                if (-not $Installed -or $item.PSIsContainer) { throw 'Source contains a reserved installation receipt.' }
                continue
            }
            if ($item.PSIsContainer) { $directories.Add($relative); $pending.Enqueue($item.FullName) }
            else { $files.Add([pscustomobject]@{ path = $relative; sha256 = Get-InstallHash $item.FullName }) }
        }
    }
    return [pscustomobject]@{ files = @($files | Sort-Object path); directories = @($directories | Sort-Object) }
}

function Assert-InstallProperties {
    param($Value, [string[]]$Names)
    if ($Value -isnot [pscustomobject]) { throw 'Invalid receipt object.' }
    $actual = @($Value.PSObject.Properties.Name)
    if ($actual.Count -ne $Names.Count -or @($actual | Where-Object { $_ -cnotin $Names }).Count -ne 0) { throw 'Invalid receipt fields.' }
}

function Test-InstallInventoryEqual {
    param($Left, $Right)
    # Exact path case and complete empty-directory inventory are part of ownership.
    return (ConvertTo-Json -InputObject $Left -Depth 10 -Compress) -ceq (ConvertTo-Json -InputObject $Right -Depth 10 -Compress)
}

function Read-InstallReceipt {
    param([string]$Target)
    Assert-InstallSafePath $Target
    $path = Join-Path $Target $script:ReceiptName
    Assert-InstallSafePath $path
    if (-not [IO.File]::Exists($path)) { throw "Unmanaged package: receipt is missing at $Target" }
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    $text = $utf8.GetString([IO.File]::ReadAllBytes($path))
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xfeff) { $text = $text.Substring(1) }
    $receipt = ConvertFrom-Json -InputObject $text -ErrorAction Stop
    Assert-InstallProperties $receipt @('schemaVersion', 'tool', 'name', 'version', 'source', 'files', 'directories')
    if ($receipt.schemaVersion -isnot [int] -or $receipt.schemaVersion -ne 1 -or $receipt.tool -cne 'MailLens' -or $receipt.name -cne 'mail-lens' -or $receipt.version -isnot [string] -or $receipt.version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$') { throw 'Unsupported or invalid installation receipt.' }
    Assert-InstallProperties $receipt.source @('state', 'commit', 'dirty')
    if ($receipt.source.state -ceq 'unknown') {
        if ($null -ne $receipt.source.commit -or $null -ne $receipt.source.dirty) { throw 'Invalid unknown source metadata.' }
    } elseif ($receipt.source.state -ceq 'git') {
        if ($receipt.source.commit -isnot [string] -or $receipt.source.commit -notmatch '^[0-9a-f]{40}$' -or $receipt.source.dirty -isnot [bool]) { throw 'Invalid Git source metadata.' }
    } else { throw 'Invalid source metadata state.' }
    if ($receipt.files -isnot [array] -or $receipt.directories -isnot [array]) { throw 'Receipt inventory must use arrays.' }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $receipt.files) {
        Assert-InstallProperties $entry @('path', 'sha256')
        Assert-InstallRelativePath $entry.path
        if ($entry.path -eq $script:ReceiptName -or -not $seen.Add($entry.path) -or $entry.sha256 -isnot [string] -or $entry.sha256 -cnotmatch '^[0-9a-f]{64}$') { throw 'Invalid or duplicate receipt file.' }
    }
    foreach ($directory in $receipt.directories) {
        Assert-InstallRelativePath $directory
        if (-not $seen.Add($directory)) { throw 'Duplicate receipt directory.' }
    }
    $expected = [pscustomobject]@{ files = @($receipt.files | Sort-Object path); directories = @($receipt.directories | Sort-Object) }
    $actual = Get-InstallInventory $Target -Installed
    if (-not (Test-InstallInventoryEqual $expected $actual)) { throw "Managed package has missing, extra or modified files/directories: $Target" }
    if (@($receipt.files | Where-Object { $_.path -ceq 'SKILL.md' }).Count -ne 1) { throw 'Receipt does not own SKILL.md.' }
    return $receipt
}

function Get-InstallSource {
    param([string]$SourcePath)
    Assert-InstallSafePath $SourcePath
    $package = Join-Path $SourcePath 'skills\mail-lens'
    $versionPath = Join-Path $SourcePath 'VERSION'
    Assert-InstallSafePath $versionPath
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    $version = $utf8.GetString([IO.File]::ReadAllBytes($versionPath))
    if ($version.Length -gt 0 -and $version[0] -eq [char]0xfeff) { $version = $version.Substring(1) }
    $version = $version.Trim()
    if ($version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$') { throw 'Invalid root VERSION.' }
    $inventory = Get-InstallInventory $package
    if (@($inventory.files | Where-Object { $_.path -ceq 'SKILL.md' }).Count -ne 1) { throw 'Source package is missing SKILL.md.' }
    $metadata = [pscustomobject]@{ state = 'unknown'; commit = $null; dirty = $null }
    if (Get-Command git -CommandType Application -ErrorAction SilentlyContinue) {
        # Git is optional metadata only; archives install without it.
        $oldPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $commit = @(& git -C $SourcePath rev-parse --verify HEAD 2>$null)
            $commitExit = $LASTEXITCODE
            $status = @(& git -C $SourcePath status --porcelain --untracked-files=normal 2>$null)
            if ($commitExit -eq 0 -and $LASTEXITCODE -eq 0 -and $commit.Count -eq 1 -and $commit[0] -match '^[0-9a-f]{40}$') { $metadata = [pscustomobject]@{ state = 'git'; commit = [string]$commit[0]; dirty = ($status.Count -gt 0) } }
        } catch { } finally { $ErrorActionPreference = $oldPreference }
    }
    return [pscustomobject]@{ schemaVersion = 1; tool = 'MailLens'; name = 'mail-lens'; version = $version; source = $metadata; files = @($inventory.files); directories = @($inventory.directories) }
}

function Resolve-InstallScope {
    param([switch]$User, [string]$Project)
    if ($User -and -not [string]::IsNullOrEmpty($Project)) { throw '-User and -Project are mutually exclusive.' }
    if (-not [string]::IsNullOrEmpty($Project)) {
        $projectRoot = Get-InstallAbsolutePath $Project 'Project'
        Assert-InstallSafePath $projectRoot
        if (-not [IO.Directory]::Exists($projectRoot)) { throw 'Project must be an existing directory.' }
        $root = Join-Path $projectRoot '.opencode'
        $scope = 'project'
    } else {
        $scope = 'user'
        if (-not [string]::IsNullOrEmpty($env:OPENCODE_CONFIG_DIR)) { $root = Get-InstallAbsolutePath $env:OPENCODE_CONFIG_DIR 'OPENCODE_CONFIG_DIR' }
        elseif (-not [string]::IsNullOrEmpty($env:XDG_CONFIG_HOME)) { $root = Join-Path (Get-InstallAbsolutePath $env:XDG_CONFIG_HOME 'XDG_CONFIG_HOME') 'opencode' }
        else {
            $userHome = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
            $root = Join-Path (Get-InstallAbsolutePath $userHome 'Windows user home') '.config\opencode'
        }
    }
    $target = Join-Path $root 'skills\mail-lens'
    Assert-InstallSafePath $target
    return [pscustomobject]@{ scope = $scope; root = $root; target = $target; lock = Join-Path $root '.mail-lens.lock' }
}

function Get-InstallTarget {
    param([string]$Target)
    Assert-InstallSafePath $Target
    if (-not (Test-Path -LiteralPath $Target)) { return $null }
    if (-not [IO.Directory]::Exists($Target)) { throw "A same-name file conflicts with the package: $Target" }
    return Read-InstallReceipt $Target
}

function Assert-InstallNoResidue {
    param([string]$Root)
    Assert-InstallSafePath $Root
    if (-not [IO.Directory]::Exists($Root)) { return }
    $residue = @(Get-ChildItem -LiteralPath $Root -Force -ErrorAction Stop | Where-Object { $_.Name -like '.mail-lens-transaction-*' })
    if ($residue.Count -gt 0) { throw ('Unfinished transaction residue; preserve and inspect recovery paths: ' + (($residue.FullName) -join ', ')) }
}

function Invoke-InstallPhase {
    param([string]$Phase, [scriptblock]$PhaseHook, [scriptblock]$CancellationRequested)
    # Private module seam for deterministic fault tests; never exposed by the CLI.
    if ($null -ne $PhaseHook) { & $PhaseHook $Phase }
    if ($null -ne $CancellationRequested -and (& $CancellationRequested)) { throw (New-Object OperationCanceledException('Installation cancelled.')) }
}

function Invoke-MailLensInstall {
    [CmdletBinding()]
    param(
        [ValidateSet('Install', 'Status', 'Update', 'Uninstall')][string]$Action = 'Install',
        [switch]$User, [string]$Project, [string]$SourcePath = $script:InstallerDefaultSource,
        [scriptblock]$CancellationRequested
    )
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or $PSVersionTable.PSVersion -lt [version]'5.1') { throw 'MailLens installation requires native Windows PowerShell 5.1 or newer.' }
    $selection = Resolve-InstallScope -User:$User -Project $Project
    $sourceRoot = Get-InstallAbsolutePath $SourcePath 'SourcePath'
    Assert-InstallSafePath $sourceRoot
    $package = Join-Path $sourceRoot 'skills\mail-lens'
    if ((Test-InstallContained $package $selection.target) -or (Test-InstallContained $selection.target $package)) { throw 'Source and target package paths must not overlap.' }
    Write-Output ("Scope: $($selection.scope); target: $($selection.target)")
    # Status reports conflicts and residue through the same fail-closed validators.
    Assert-InstallNoResidue $selection.root
    $original = Get-InstallTarget $selection.target
    if ($Action -eq 'Status') {
        if ($null -eq $original) { Write-Output 'Status: absent.' }
        else {
            Write-Output ("Status: managed, integrity verified; version $($original.version); source metadata $($original.source.state).")
            if (-not [IO.File]::Exists((Join-Path $sourceRoot 'VERSION')) -or -not [IO.Directory]::Exists($package)) { Write-Output 'Selected source is unavailable; installed ownership/integrity remains verified.' }
        }
        return
    }
    if ($Action -eq 'Uninstall' -and $null -eq $original) { Write-Output 'Uninstall: absent; unchanged.'; return }
    if ($Action -eq 'Update' -and $null -eq $original) { throw 'Update requires an unchanged managed installation.' }
    $source = $null
    if ($Action -ne 'Uninstall') {
        $source = Get-InstallSource $sourceRoot
        if ($null -ne $original) {
            $same = $source.version -ceq $original.version -and (Test-InstallInventoryEqual ([pscustomobject]@{files = $source.files; directories = $source.directories}) ([pscustomobject]@{files = $original.files; directories = $original.directories}))
            if ($same) { Write-Output ("${Action}: unchanged; receipt and package were not rewritten."); return }
            if ($Action -eq 'Install') { throw 'Installed package differs from selected source. Use explicit -Action Update.' }
        }
    }
    Assert-InstallSafePath $selection.lock
    $lockHandle = $null
    $transaction = $null
    $backup = $null
    $movedOriginal = $false
    $published = $false
    $committed = $false
    $recoveryFailed = $false
    try {
        [void][IO.Directory]::CreateDirectory($selection.root)
        try { $lockHandle = [IO.File]::Open($selection.lock, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch { throw "Another operation holds the target lock, or the lock cannot be opened: $($selection.lock)" }
        Assert-InstallNoResidue $selection.root
        $current = Get-InstallTarget $selection.target
        if (($null -eq $current) -ne ($null -eq $original) -or ($null -ne $current -and (ConvertTo-Json $current -Depth 20 -Compress) -cne (ConvertTo-Json $original -Depth 20 -Compress))) { throw 'Target changed before locking; retry after inspection.' }
        $transaction = Join-Path $selection.root ('.mail-lens-transaction-' + [Guid]::NewGuid().ToString('N'))
        Assert-InstallSafePath $transaction
        [void][IO.Directory]::CreateDirectory($transaction)
        $staged = Join-Path $transaction 'new-package'
        $backup = Join-Path $transaction 'old-package'
        if ($Action -ne 'Uninstall') {
            $fresh = Get-InstallSource $sourceRoot
            if ((ConvertTo-Json $source -Depth 20 -Compress) -cne (ConvertTo-Json $fresh -Depth 20 -Compress)) { throw 'Source changed before staging.' }
            [void][IO.Directory]::CreateDirectory($staged)
            foreach ($directory in $source.directories) { [void][IO.Directory]::CreateDirectory((Join-Path $staged $directory)) }
            foreach ($file in $source.files) {
                Invoke-InstallPhase 'stage-write' $script:InstallPhaseHook $CancellationRequested
                [IO.File]::Copy((Join-Path $package $file.path), (Join-Path $staged $file.path), $false)
            }
            $utf8 = New-Object Text.UTF8Encoding($false, $true)
            [IO.File]::WriteAllText((Join-Path $staged $script:ReceiptName), (ConvertTo-Json -InputObject $source -Depth 20) + "`n", $utf8)
            [void](Read-InstallReceipt $staged)
        }
        Invoke-InstallPhase 'before-swap' $script:InstallPhaseHook $CancellationRequested
        Assert-InstallSafePath $selection.target
        Assert-InstallSafePath $staged
        Assert-InstallSafePath $backup
        $current = Get-InstallTarget $selection.target
        if (($null -eq $current) -ne ($null -eq $original) -or ($null -ne $current -and (ConvertTo-Json $current -Depth 20 -Compress) -cne (ConvertTo-Json $original -Depth 20 -Compress))) { throw 'Target changed before publication.' }
        if ($Action -ne 'Uninstall') {
            $fresh = Get-InstallSource $sourceRoot
            if ((ConvertTo-Json $source -Depth 20 -Compress) -cne (ConvertTo-Json $fresh -Depth 20 -Compress)) { throw 'Source changed during staging.' }
            [void](Read-InstallReceipt $staged)
        }
        if ($null -ne $original) {
            Invoke-InstallPhase 'target-move' $script:InstallPhaseHook $CancellationRequested
            [IO.Directory]::Move($selection.target, $backup)
            $movedOriginal = $true
        }
        if ($Action -ne 'Uninstall') {
            [void][IO.Directory]::CreateDirectory((Split-Path $selection.target -Parent))
            Invoke-InstallPhase 'publish' $script:InstallPhaseHook $CancellationRequested
            [IO.Directory]::Move($staged, $selection.target)
            $published = $true
        }
        Invoke-InstallPhase 'before-commit' $script:InstallPhaseHook $CancellationRequested
        $committed = $true
        Write-Output ("${Action}: completed at $($selection.target).")
    } catch {
        $failure = $_
        if (-not $committed -and ($movedOriginal -or $published)) {
            try {
                Invoke-InstallPhase 'restore' $script:InstallPhaseHook $null
                if ($published) {
                    # Never remove unexplained edits made after the complete directory swap.
                    [void](Read-InstallReceipt $selection.target)
                    [IO.Directory]::Move($selection.target, $staged)
                }
                if ($movedOriginal) { [IO.Directory]::Move($backup, $selection.target) }
            } catch {
                $recoveryFailed = $true
                if ($movedOriginal) { $recovery = "full previous package backup: $backup" }
                else { $recovery = "complete published package: $($selection.target)" }
                throw ("Operation failed and restoration failed. Preserve transaction: $transaction ; $recovery . Inspect target before manual recovery.")
            }
        }
        throw $failure
    } finally {
        if ($null -ne $transaction -and -not $recoveryFailed -and [IO.Directory]::Exists($transaction)) {
            try {
                Invoke-InstallPhase 'cleanup' $script:InstallPhaseHook $null
                # Validate every tree before bounded cleanup; never follow injected links.
                Assert-InstallSafePath $transaction
                foreach ($child in @(Get-ChildItem -LiteralPath $transaction -Force)) {
                    if (-not $child.PSIsContainer -or $child.Name -notin @('new-package', 'old-package')) { throw 'Unexpected transaction content.' }
                    [void](Get-InstallInventory $child.FullName -Installed)
                }
                [IO.Directory]::Delete($transaction, $true)
            } catch {
                Write-Warning ("Owned transaction cleanup failed; committed=$committed; residue retained: $transaction")
            }
        }
        if ($null -ne $lockHandle) {
            $lockHandle.Dispose()
            try { [IO.File]::Delete($selection.lock) }
            catch { Write-Warning ("Lock cleanup failed; inspect residue: $($selection.lock)") }
        }
    }
}


# Dot-sourcing is reserved for development tests of the same functions. The CLI
# exposes no fault, force, recovery, policy or configuration override switch.
if ($MyInvocation.InvocationName -ne '.') {
    if (-not ('MailLens.InstallCancellation' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
namespace MailLens {
    public static class InstallCancellation {
        public static volatile bool Requested;
        private static void OnCancel(object sender, ConsoleCancelEventArgs args) {
            Requested = true; args.Cancel = true;
        }
        public static void Start() { Requested = false; Console.CancelKeyPress += OnCancel; }
        public static void Stop() { Console.CancelKeyPress -= OnCancel; }
    }
}
'@
    }
    $code = 1
    try {
        [MailLens.InstallCancellation]::Start()
        Invoke-MailLensInstall -Action $Action -User:$User -Project $Project -SourcePath $SourcePath -CancellationRequested { [MailLens.InstallCancellation]::Requested }
        $code = 0
    } catch { [Console]::Error.WriteLine('MailLens installer: ' + $_.Exception.Message) }
    finally { [MailLens.InstallCancellation]::Stop() }
    exit $code
}
