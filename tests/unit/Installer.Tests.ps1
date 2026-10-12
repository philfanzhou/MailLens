BeforeAll {
    . (Join-Path $PSScriptRoot '../../scripts/install.ps1')
    $script:Repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    $script:Utf8 = New-Object Text.UTF8Encoding($false, $true)
    function Write-TestText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:Utf8) }
    function Get-TestSnapshot([string]$Root) {
        if (-not [IO.Directory]::Exists($Root)) { return 'absent' }
        $rows = @(Get-ChildItem -LiteralPath $Root -Force -Recurse | Sort-Object FullName | ForEach-Object {
            $relative = $_.FullName.Substring($Root.Length)
            if ($_.PSIsContainer) { 'D:' + $relative }
            else { 'F:' + $relative + ':' + (Get-InstallHash $_.FullName) + ':' + $_.LastWriteTimeUtc.Ticks }
        })
        return $rows -join "`n"
    }
    function Write-TestReceipt($Receipt) {
        Write-TestText (Join-Path $script:Target '.mail-lens-install.json') (ConvertTo-Json $Receipt -Depth 20)
    }
    function Install-TestPackage { Invoke-MailLensInstall -Project $script:Project -SourcePath $script:Source | Out-Null }
    function Change-TestSource { Write-TestText (Join-Path $script:Source 'skills/mail-lens/addition.txt') 'synthetic newer content' }
    function Invoke-TestUpdate { Invoke-MailLensInstall -Action Update -Project $script:Project -SourcePath $script:Source | Out-Null }
}

Describe 'Windows installer scope and validation' {
    BeforeEach {
        $script:CaseRoot = Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-installer-' + [Guid]::NewGuid().ToString('N') + '-中文 space')
        $script:Source = Join-Path $script:CaseRoot 'fixed source'
        $script:Project = Join-Path $script:CaseRoot 'project 中文'
        $script:Config = Join-Path $script:Project '.opencode'
        $script:Target = Join-Path $script:Config 'skills/mail-lens'
        [void][IO.Directory]::CreateDirectory((Join-Path $script:Source 'skills'))
        [void][IO.Directory]::CreateDirectory($script:Project)
        Copy-Item -LiteralPath (Join-Path $script:Repository 'skills/mail-lens') -Destination (Join-Path $script:Source 'skills/mail-lens') -Recurse
        [void][IO.Directory]::CreateDirectory((Join-Path $script:Source 'skills/mail-lens/empty 中文'))
        Write-TestText (Join-Path $script:Source 'VERSION') '0.1.0'
        $script:SavedConfigDir = $env:OPENCODE_CONFIG_DIR
        $script:SavedXdg = $env:XDG_CONFIG_HOME
        $script:SavedConfigFile = $env:OPENCODE_CONFIG
        $env:OPENCODE_CONFIG_DIR = $null
        $env:XDG_CONFIG_HOME = $null
        $env:OPENCODE_CONFIG = $null
        $script:InstallPhaseHook = $null
        $script:CancelRequested = $false
    }
    AfterEach {
        $script:InstallPhaseHook = $null
        $env:OPENCODE_CONFIG_DIR = $script:SavedConfigDir
        $env:XDG_CONFIG_HOME = $script:SavedXdg
        $env:OPENCODE_CONFIG = $script:SavedConfigFile
        if ([IO.Directory]::Exists($script:CaseRoot)) { [IO.Directory]::Delete($script:CaseRoot, $true) }
    }

    It 'resolves default and User to actual Windows home, without installing there' {
        $expected = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.config/opencode/skills/mail-lens'
        (Resolve-InstallScope).target | Should -Be $expected
        (Resolve-InstallScope -User).target | Should -Be $expected
    }
    It 'resolves custom config before XDG and ignores config-file override' {
        $env:OPENCODE_CONFIG_DIR = Join-Path $script:CaseRoot 'custom 中文'
        $env:XDG_CONFIG_HOME = Join-Path $script:CaseRoot 'xdg space'
        $env:OPENCODE_CONFIG = 'relative-config-file.json'
        (Resolve-InstallScope).target | Should -Be (Join-Path $env:OPENCODE_CONFIG_DIR 'skills/mail-lens')
        $env:OPENCODE_CONFIG_DIR = $null
        (Resolve-InstallScope).target | Should -Be (Join-Path $env:XDG_CONFIG_HOME 'opencode/skills/mail-lens')
    }
    It 'installs in isolated default/global and explicit User custom roots' {
        $env:OPENCODE_CONFIG_DIR = Join-Path $script:CaseRoot 'global config'
        Invoke-MailLensInstall -SourcePath $script:Source | Out-Null
        $globalTarget = (Resolve-InstallScope).target
        [void](Read-InstallReceipt $globalTarget)
        $snapshot = Get-TestSnapshot $globalTarget
        Invoke-MailLensInstall -User -SourcePath $script:Source | Out-Null
        (Get-TestSnapshot $globalTarget) | Should -BeExactly $snapshot
    }
    It 'rejects invalid selected override <Name> without creating config' -ForEach @(@{Name='OPENCODE_CONFIG_DIR'}, @{Name='XDG_CONFIG_HOME'}) {
        [Environment]::SetEnvironmentVariable($Name, 'relative/skills', 'Process')
        { Invoke-MailLensInstall -SourcePath $script:Source } | Should -Throw '*absolute Windows path*'
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }
    It 'rejects conflicting User and Project' {
        { Invoke-MailLensInstall -User -Project $script:Project -SourcePath $script:Source } | Should -Throw '*mutually exclusive*'
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }
    It 'rejects relative, missing, invalid Project/source/action values' {
        { Invoke-MailLensInstall -Project 'relative' -SourcePath $script:Source } | Should -Throw
        { Invoke-MailLensInstall -Project (Join-Path $script:CaseRoot 'missing') -SourcePath $script:Source } | Should -Throw
        { Invoke-MailLensInstall -Project $script:Project -SourcePath 'relative' } | Should -Throw
        { Invoke-MailLensInstall -Action Force -Project $script:Project -SourcePath $script:Source } | Should -Throw
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }
    It 'uses separator-aware case-insensitive overlap checks, excluding prefix traps' {
        Test-InstallContained 'C:\path\mail-lens' 'c:\PATH\MAIL-LENS\scripts' | Should -BeTrue
        Test-InstallContained 'C:\path\mail-lens' 'C:\path\mail-lens-extra' | Should -BeFalse
        Test-InstallContained 'C:\path\mail-lens' 'C:\path\mail' | Should -BeFalse
    }
    It 'rejects direct and nested source-target overlap before writes' {
        $env:OPENCODE_CONFIG_DIR = $script:Source
        { Invoke-MailLensInstall -SourcePath $script:Source } | Should -Throw '*overlap*'
        $env:OPENCODE_CONFIG_DIR = Join-Path $script:Source 'skills/mail-lens/nested'
        { Invoke-MailLensInstall -SourcePath $script:Source } | Should -Throw '*overlap*'
        [IO.File]::Exists((Join-Path $script:Source '.mail-lens.lock')) | Should -BeFalse
    }
    It 'preserves an unmanaged directory or same-name file' -ForEach @(@{Kind='directory'}, @{Kind='file'}) {
        [void][IO.Directory]::CreateDirectory((Split-Path $script:Target -Parent))
        if ($Kind -eq 'directory') { [void][IO.Directory]::CreateDirectory($script:Target); Write-TestText (Join-Path $script:Target 'sentinel') 'user data' }
        else { Write-TestText $script:Target 'user data' }
        $before = Get-TestSnapshot $script:Project
        { Install-TestPackage } | Should -Throw
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
    }

    It 'copies every byte, includes empty directories and writes strict BOM-free schema 1 receipt' {
        Install-TestPackage
        $receipt = Read-InstallReceipt $script:Target
        $receipt.version | Should -Be '0.1.0'
        $receipt.source.state | Should -Be 'unknown'
        $receipt.source.commit | Should -BeNullOrEmpty
        $receipt.source.dirty | Should -BeNullOrEmpty
        $receipt.files.Count | Should -Be 5
        $receipt.directories | Should -Contain 'empty 中文'
        foreach ($file in $receipt.files) { (Get-InstallHash (Join-Path $script:Target $file.path)) | Should -Be (Get-InstallHash (Join-Path $script:Source ('skills/mail-lens/' + $file.path))) }
        $bytes = [IO.File]::ReadAllBytes((Join-Path $script:Target '.mail-lens-install.json'))
        $bytes[0] | Should -Be 123
        $script:Utf8.GetString($bytes) | Should -Match '"schemaVersion":\s+1'
    }
    It 'does not rewrite status, identical install or identical update including receipt timestamps' {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Project
        foreach ($action in @('Status', 'Install', 'Update')) { Invoke-MailLensInstall -Action $action -Project $script:Project -SourcePath $script:Source | Out-Null; (Get-TestSnapshot $script:Project) | Should -BeExactly $before }
        [IO.File]::Exists((Join-Path $script:Config '.mail-lens.lock')) | Should -BeFalse
    }
    It 'reports absent status/repeated uninstall without creating parent paths' {
        $before = Get-TestSnapshot $script:Project
        foreach ($action in @('Status', 'Uninstall', 'Uninstall')) { Invoke-MailLensInstall -Action $action -Project $script:Project -SourcePath (Join-Path $script:CaseRoot 'gone') | Out-Null }
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }
    It 'uses installed launcher/module/references after source removal and source-free retained CLI status/uninstall' {
        Install-TestPackage
        $retained = Join-Path $script:CaseRoot 'retained-install.ps1'
        [IO.File]::Copy((Join-Path $script:Repository 'scripts/install.ps1'), $retained)
        [IO.Directory]::Delete($script:Source, $true)
        Import-Module (Join-Path $script:Target 'scripts/MailLens.Outlook.psm1') -Force
        (Get-Command Invoke-MailLensFileRequest).Name | Should -Be 'Invoke-MailLensFileRequest'
        $tokens = $null; $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Target 'scripts/outlook.ps1'), [ref]$tokens, [ref]$errors)
        $errors.Count | Should -Be 0
        foreach ($reference in @('references/commands.md', 'references/workflow.md')) { [IO.File]::Exists((Join-Path $script:Target $reference)) | Should -BeTrue }
        $before = Get-TestSnapshot $script:Project
        $report = & powershell.exe -NoProfile -File $retained -Action Status -Project $script:Project -SourcePath $script:Source 2>&1
        $LASTEXITCODE | Should -Be 0
        ($report -join "`n") | Should -Match 'source is unavailable'
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
        & powershell.exe -NoProfile -File $retained -Action Uninstall -Project $script:Project -SourcePath $script:Source | Out-Null
        $LASTEXITCODE | Should -Be 0
        [IO.Directory]::Exists($script:Target) | Should -BeFalse
        Remove-Module MailLens.Outlook
    }
    It 'requires explicit Update and permits deliberate older-content rollback' {
        Install-TestPackage
        $old = Get-TestSnapshot $script:Target
        Change-TestSource
        $before = Get-TestSnapshot $script:Project
        { Install-TestPackage } | Should -Throw '*explicit*Update*'
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
        Invoke-TestUpdate
        [IO.File]::Exists((Join-Path $script:Target 'addition.txt')) | Should -BeTrue
        [IO.File]::Delete((Join-Path $script:Source 'skills/mail-lens/addition.txt'))
        Invoke-TestUpdate
        [IO.File]::Exists((Join-Path $script:Target 'addition.txt')) | Should -BeFalse
        [void](Read-InstallReceipt $script:Target)
        (Get-InstallInventory $script:Target -Installed).files.Count | Should -Be 5
    }
    It 'takes version only from root VERSION and updates same bytes with changed version' {
        Install-TestPackage
        Write-TestText (Join-Path $script:Source 'VERSION') '0.0.9'
        Invoke-TestUpdate
        (Read-InstallReceipt $script:Target).version | Should -Be '0.0.9'
    }
    It 'rejects absent update without config writes' {
        { Invoke-TestUpdate } | Should -Throw '*requires*managed*'
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }
    It 'uninstalls only selected package preserving configuration, another skill and other scopes' {
        Install-TestPackage
        Write-TestText (Join-Path $script:Config 'opencode.json') '{"synthetic":true}'
        $other = Join-Path $script:Config 'skills/other'
        [void][IO.Directory]::CreateDirectory($other)
        Write-TestText (Join-Path $other 'SKILL.md') 'other synthetic skill'
        $globalRoot = Join-Path $script:CaseRoot 'global'
        $env:OPENCODE_CONFIG_DIR = $globalRoot
        Invoke-MailLensInstall -User -SourcePath $script:Source | Out-Null
        $globalBefore = Get-TestSnapshot $globalRoot
        Invoke-MailLensInstall -Action Uninstall -Project $script:Project -SourcePath $script:Source | Out-Null
        [IO.Directory]::Exists($script:Target) | Should -BeFalse
        [IO.File]::ReadAllText((Join-Path $script:Config 'opencode.json')) | Should -Be '{"synthetic":true}'
        [IO.File]::ReadAllText((Join-Path $other 'SKILL.md')) | Should -Be 'other synthetic skill'
        (Get-TestSnapshot $globalRoot) | Should -BeExactly $globalBefore
    }

    It 'rejects and preserves target <Tamper> for status/install/update/uninstall' -ForEach @(@{Tamper='edited-file'}, @{Tamper='missing-file'}, @{Tamper='extra-file'}, @{Tamper='extra-dir'}, @{Tamper='missing-empty-dir'}, @{Tamper='corrupt-receipt'}) {
        Install-TestPackage
        switch ($Tamper) {
            'edited-file' { Write-TestText (Join-Path $script:Target 'SKILL.md') 'user edited' }
            'missing-file' { [IO.File]::Delete((Join-Path $script:Target 'SKILL.md')) }
            'extra-file' { Write-TestText (Join-Path $script:Target 'user.txt') 'user data' }
            'extra-dir' { [void][IO.Directory]::CreateDirectory((Join-Path $script:Target 'user dir')) }
            'missing-empty-dir' { [IO.Directory]::Delete((Join-Path $script:Target 'empty 中文')) }
            'corrupt-receipt' { Write-TestText (Join-Path $script:Target '.mail-lens-install.json') 'not json' }
        }
        $before = Get-TestSnapshot $script:Project
        foreach ($action in @('Status', 'Install', 'Update', 'Uninstall')) { { Invoke-MailLensInstall -Action $action -Project $script:Project -SourcePath $script:Source } | Should -Throw; (Get-TestSnapshot $script:Project) | Should -BeExactly $before }
    }
    It 'rejects unsafe/invalid receipt <Tamper> before mutation' -ForEach @(@{Tamper='traversal'}, @{Tamper='absolute'}, @{Tamper='duplicate'}, @{Tamper='case-collision'}, @{Tamper='wrong-tool'}, @{Tamper='wrong-name'}, @{Tamper='schema'}, @{Tamper='hash'}, @{Tamper='unknown-field'}, @{Tamper='source'}, @{Tamper='stream'}, @{Tamper='reserved'}) {
        Install-TestPackage
        $receipt = Read-InstallReceipt $script:Target
        switch ($Tamper) {
            'traversal' { $receipt.files[0].path = '../outside.txt' }
            'absolute' { $receipt.files[0].path = 'C:/outside.txt' }
            'duplicate' { $receipt.files += $receipt.files[0] }
            'case-collision' { $receipt.files += [pscustomobject]@{path=$receipt.files[0].path.ToUpperInvariant();sha256=$receipt.files[0].sha256} }
            'wrong-tool' { $receipt.tool = 'Other' }
            'wrong-name' { $receipt.name = 'other' }
            'schema' { $receipt.schemaVersion = 2 }
            'hash' { $receipt.files[0].sha256 = 'bad' }
            'unknown-field' { $receipt | Add-Member secret 'unexpected' }
            'source' { $receipt.source.state = 'clean' }
            'stream' { $receipt.files[0].path = 'file.txt:stream' }
            'reserved' { $receipt.files[0].path = 'CON.txt' }
        }
        Write-TestReceipt $receipt
        $before = Get-TestSnapshot $script:Project
        { Invoke-MailLensInstall -Action Uninstall -Project $script:Project -SourcePath $script:Source } | Should -Throw
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
    }
    It 'accepts optional receipt UTF-8 BOM but rejects malformed UTF-8' {
        Install-TestPackage
        $path = Join-Path $script:Target '.mail-lens-install.json'
        $bytes = [IO.File]::ReadAllBytes($path)
        [IO.File]::WriteAllBytes($path, ([byte[]]@(239,187,191) + $bytes))
        [void](Read-InstallReceipt $script:Target)
        [IO.File]::WriteAllBytes($path, [byte[]]@(123,255,125))
        { Read-InstallReceipt $script:Target } | Should -Throw
    }
    It 'rejects reserved source receipt and invalid VERSION without creating target' {
        Write-TestText (Join-Path $script:Source 'skills/mail-lens/.mail-lens-install.json') '{}'
        { Install-TestPackage } | Should -Throw '*reserved*'
        [IO.File]::Delete((Join-Path $script:Source 'skills/mail-lens/.mail-lens-install.json'))
        Write-TestText (Join-Path $script:Source 'VERSION') 'latest'
        { Install-TestPackage } | Should -Throw '*VERSION*'
        [IO.Directory]::Exists($script:Config) | Should -BeFalse
    }

    It 'restores original bytes/receipt/timestamps after <Phase> failure on update' -ForEach @(@{Phase='stage-write'}, @{Phase='target-move'}, @{Phase='publish'}, @{Phase='before-commit'}) {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        Change-TestSource
        $script:FaultPhase = $Phase
        $script:InstallPhaseHook = { param($p) if ($p -eq $script:FaultPhase) { throw 'Injected phase failure' } }
        { Invoke-TestUpdate } | Should -Throw '*Injected phase failure*'
        (Get-TestSnapshot $script:Target) | Should -BeExactly $before
        @(Get-ChildItem $script:Config -Force | Where-Object Name -like '.mail-lens-transaction-*').Count | Should -Be 0
        [IO.File]::Exists((Join-Path $script:Config '.mail-lens.lock')) | Should -BeFalse
    }
    It 'leaves an absent original absent after failed initial publication' {
        $script:InstallPhaseHook = { param($p) if ($p -eq 'before-commit') { throw 'Injected initial failure' } }
        { Install-TestPackage } | Should -Throw
        [IO.Directory]::Exists($script:Target) | Should -BeFalse
    }
    It 'restores uninstall before commit after cooperative cancellation' {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        $script:InstallPhaseHook = { param($p) if ($p -eq 'before-commit') { $script:CancelRequested = $true } }
        { Invoke-MailLensInstall -Action Uninstall -Project $script:Project -SourcePath $script:Source -CancellationRequested { $script:CancelRequested } } | Should -Throw '*cancelled*'
        (Get-TestSnapshot $script:Target) | Should -BeExactly $before
    }
    It 'restores update after cooperative cancellation between old and new moves' {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -eq 'publish') { $script:CancelRequested = $true } }
        { Invoke-MailLensInstall -Action Update -Project $script:Project -SourcePath $script:Source -CancellationRequested { $script:CancelRequested } } | Should -Throw '*cancelled*'
        (Get-TestSnapshot $script:Target) | Should -BeExactly $before
    }
    It 'preserves full old backup and reports exact path when restoration fails' {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -in @('before-commit', 'restore')) { throw 'Injected restore fault' } }
        $failure = $null
        try { Invoke-TestUpdate } catch { $failure = $_.Exception.Message }
        $failure | Should -Match 'restoration failed'
        $transaction = @(Get-ChildItem $script:Config -Force | Where-Object Name -like '.mail-lens-transaction-*')[0].FullName
        $backup = Join-Path $transaction 'old-package'
        $failure.Contains($backup) | Should -BeTrue
        (Get-TestSnapshot $backup) | Should -BeExactly $before
        $script:InstallPhaseHook = $null
        { Invoke-TestUpdate } | Should -Throw '*residue*'
        (Get-TestSnapshot $backup) | Should -BeExactly $before
    }
    It 'keeps committed <Operation> and reports cleanup residue distinctly' -ForEach @(@{Operation='Update'}, @{Operation='Uninstall'}) {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -eq 'cleanup') { throw 'Injected cleanup fault' } }
        $report = @(Invoke-MailLensInstall -Action $Operation -Project $script:Project -SourcePath $script:Source 3>&1)
        ($report -join "`n") | Should -Match 'completed'
        ($report -join "`n") | Should -Match 'committed=True; residue retained:'
        $transaction = @(Get-ChildItem $script:Config -Force | Where-Object Name -like '.mail-lens-transaction-*')[0].FullName
        (Get-TestSnapshot (Join-Path $transaction 'old-package')) | Should -BeExactly $before
        if ($Operation -eq 'Update') { [void](Read-InstallReceipt $script:Target); [IO.File]::Exists((Join-Path $script:Target 'addition.txt')) | Should -BeTrue }
        else { [IO.Directory]::Exists($script:Target) | Should -BeFalse }
    }
    It 'detects source changes during staging before publication' {
        Install-TestPackage
        $before = Get-TestSnapshot $script:Target
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -eq 'before-swap') { Write-TestText (Join-Path $script:Source 'skills/mail-lens/addition.txt') 'changed while staging' } }
        { Invoke-TestUpdate } | Should -Throw '*Source changed during staging*'
        (Get-TestSnapshot $script:Target) | Should -BeExactly $before
    }
    It 'detects local edits made after locking and preserves them' {
        Install-TestPackage
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -eq 'before-swap') { Write-TestText (Join-Path $script:Target 'SKILL.md') 'concurrent user edit' } }
        { Invoke-TestUpdate } | Should -Throw '*modified*'
        [IO.File]::ReadAllText((Join-Path $script:Target 'SKILL.md')) | Should -Be 'concurrent user edit'
    }
    It 'stages outside skill discovery with original still complete until swap' {
        Install-TestPackage
        Change-TestSource
        $script:InstallPhaseHook = { param($p) if ($p -eq 'before-swap') {
            [void](Read-InstallReceipt $script:Target)
            $transaction = @(Get-ChildItem $script:Config -Force | Where-Object Name -like '.mail-lens-transaction-*')[0].FullName
            (Test-InstallContained (Join-Path $script:Config 'skills') $transaction) | Should -BeFalse
            [void](Read-InstallReceipt (Join-Path $transaction 'new-package'))
        } }
        Invoke-TestUpdate
    }
    It 'detects unfinished residue without removing any owned prior backup' {
        $residue = Join-Path $script:Config '.mail-lens-transaction-synthetic'
        [void][IO.Directory]::CreateDirectory($residue)
        Write-TestText (Join-Path $residue 'sentinel') 'preserve'
        $before = Get-TestSnapshot $script:Project
        { Install-TestPackage } | Should -Throw '*recovery paths*'
        (Get-TestSnapshot $script:Project) | Should -BeExactly $before
    }
    It 'refuses genuine second-process target lock and permits an independent target' {
        Install-TestPackage
        Change-TestSource
        $before = Get-TestSnapshot $script:Target
        $ready = Join-Path $script:CaseRoot 'lock-ready'
        $release = Join-Path $script:CaseRoot 'lock-release'
        $holder = Join-Path $script:CaseRoot 'hold.ps1'
        $lock = Join-Path $script:Config '.mail-lens.lock'
        $content = @'
param($Lock, $Ready, $Release)
$handle = [IO.File]::Open($Lock, 'OpenOrCreate', 'ReadWrite', 'None')
try {
    [IO.File]::WriteAllText($Ready, 'ready')
    $limit = [DateTime]::UtcNow.AddSeconds(30)
    while (-not [IO.File]::Exists($Release) -and [DateTime]::UtcNow -lt $limit) { Start-Sleep -Milliseconds 100 }
} finally { $handle.Dispose() }
'@
        Write-TestText $holder $content
        $process = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-File', ('"' + $holder + '"'), '-Lock', ('"' + $lock + '"'), '-Ready', ('"' + $ready + '"'), '-Release', ('"' + $release + '"')) -PassThru -WindowStyle Hidden
        try {
            $limit = [DateTime]::UtcNow.AddSeconds(10)
            while (-not [IO.File]::Exists($ready) -and [DateTime]::UtcNow -lt $limit) { Start-Sleep -Milliseconds 100 }
            [IO.File]::Exists($ready) | Should -BeTrue
            { Invoke-TestUpdate } | Should -Throw '*target lock*'
            (Get-TestSnapshot $script:Target) | Should -BeExactly $before
            $other = Join-Path $script:CaseRoot 'independent project'
            [void][IO.Directory]::CreateDirectory($other)
            Invoke-MailLensInstall -Project $other -SourcePath $script:Source | Out-Null
            [void](Read-InstallReceipt (Join-Path $other '.opencode/skills/mail-lens'))
        } finally { Write-TestText $release 'release'; $process.WaitForExit(15000) | Out-Null; if (-not $process.HasExited) { $process.Kill() }; [IO.File]::Delete($lock) }
        Invoke-TestUpdate
    }
    It 'rejects native junction in <Location> without touching linked content' -ForEach @(@{Location='source-tree'}, @{Location='target-tree'}, @{Location='ancestor'}, @{Location='lock'}) {
        $outside = Join-Path $script:CaseRoot 'outside'
        [void][IO.Directory]::CreateDirectory($outside)
        Write-TestText (Join-Path $outside 'sentinel') 'unrelated'
        $link = $null
        if ($Location -eq 'source-tree') { $link = Join-Path $script:Source 'skills/mail-lens/link' }
        elseif ($Location -eq 'target-tree') { Install-TestPackage; $link = Join-Path $script:Target 'link' }
        elseif ($Location -eq 'ancestor') { $link = $script:Config }
        else { [void][IO.Directory]::CreateDirectory($script:Config); $link = Join-Path $script:Config '.mail-lens.lock' }
        New-Item -ItemType Junction -Path $link -Target $outside -ErrorAction Stop | Out-Null
        try {
            { Install-TestPackage } | Should -Throw '*reparse*'
            [IO.File]::ReadAllText((Join-Path $outside 'sentinel')) | Should -Be 'unrelated'
        } finally { [IO.Directory]::Delete($link) }
    }
    It 'CLI uses truthful success and failure exit codes' {
        $installer = Join-Path $script:Repository 'scripts/install.ps1'
        & powershell.exe -NoProfile -File $installer -Project $script:Project -SourcePath $script:Source | Out-Null
        $LASTEXITCODE | Should -Be 0
        Write-TestText (Join-Path $script:Target 'SKILL.md') 'user data'
        $preference = $ErrorActionPreference
        try { $ErrorActionPreference = 'Continue'; & powershell.exe -NoProfile -File $installer -Action Uninstall -Project $script:Project -SourcePath $script:Source 2>&1 | Out-Null; $LASTEXITCODE | Should -Be 1 }
        finally { $ErrorActionPreference = $preference }
    }
}
