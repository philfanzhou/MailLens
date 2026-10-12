BeforeAll {
    $script:Repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    . (Join-Path $script:Repository 'scripts/ci-process.ps1')
    $script:PowerShell = (Get-Command powershell.exe -CommandType Application -ErrorAction Stop).Source
    $script:Utf8 = New-Object Text.UTF8Encoding($false, $true)
    if (-not [IO.File]::Exists($env:MAIL_LENS_OPENCODE)) { throw 'Pinned native OpenCode is required.' }
    function Write-DiscoveryText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:Utf8) }
    function Invoke-Installer([string]$Action = 'Install', [string]$Source = $script:Source, [string]$Installer = $script:Installer, [string[]]$Selection = $script:Selection) {
        $arguments = @('-NoProfile', '-STA', '-File', $Installer, '-Action', $Action, '-SourcePath', $Source) + $Selection
        Invoke-CIProcess $script:PowerShell $arguments $script:Project
    }
    function Get-DiscoveredSkills {
        $result = Invoke-CIProcess $env:MAIL_LENS_OPENCODE @('debug', 'skill') $script:Project
        if ($result.Code -ne 0) { throw "OpenCode discovery failed: $($result.Error)" }
        ConvertFrom-CISkills $result.Out
    }
    function Assert-DiscoveredPackage([string]$Target = $script:Target, [string]$Source = $script:Source) {
        $entry = [IO.File]::ReadAllText((Join-Path $Source 'skills/mail-lens/SKILL.md'), $script:Utf8).Replace("`r`n", "`n")
        $front = [regex]::Match($entry, '\A---\n.*?\n---\n', [Text.RegularExpressions.RegexOptions]::Singleline)
        $description = [regex]::Match($front.Value, '(?m)^description: (.+)$').Groups[1].Value.TrimEnd("`r")
        Assert-CITarget @(Get-DiscoveredSkills) (Join-Path $Target 'SKILL.md') $entry.Substring($front.Length) $description
        foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $Source 'skills/mail-lens') -File -Recurse)) {
            $relative = $file.FullName.Substring((Join-Path $Source 'skills/mail-lens').Length + 1)
            (Get-FileHash -LiteralPath (Join-Path $Target $relative) -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
    }
    function Get-DiscoverySnapshot([string]$Root) {
        @((Get-ChildItem -LiteralPath $Root -File -Recurse -Force | Sort-Object FullName | ForEach-Object {
            $_.FullName.Substring($Root.Length) + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash + ':' + $_.LastWriteTimeUtc.Ticks
        })) -join "`n"
    }
    function Select-DiscoveryScope([string]$Scope) {
        $script:Selection = @(); $env:OPENCODE_CONFIG_DIR = $null
        if ($Scope -eq 'default') { $script:Target = Join-Path $env:XDG_CONFIG_HOME 'opencode/skills/mail-lens' }
        elseif ($Scope -eq 'User') { $env:OPENCODE_CONFIG_DIR = Join-Path $script:CaseRoot 'custom config'; $script:Selection = @('-User'); $script:Target = Join-Path $env:OPENCODE_CONFIG_DIR 'skills/mail-lens' }
        else { $script:Selection = @('-Project', $script:Project); $script:Target = Join-Path $script:Project '.opencode/skills/mail-lens' }
    }
}
Describe 'Pinned OpenCode native isolated discovery' {
    BeforeEach {
        $script:CaseRoot = Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-discovery-' + [Guid]::NewGuid().ToString('N') + '-中文 space')
        $script:Project = Join-Path $script:CaseRoot 'project'; $script:Source = Join-Path $script:CaseRoot 'source'
        [void][IO.Directory]::CreateDirectory($script:Project); [void][IO.Directory]::CreateDirectory($script:Source)
        & git -C $script:Project init --quiet
        if ($LASTEXITCODE -ne 0) { throw 'Cannot create isolated project boundary.' }
        Copy-Item -LiteralPath (Join-Path $script:Repository 'skills') -Destination $script:Source -Recurse
        Copy-Item -LiteralPath (Join-Path $script:Repository 'VERSION') -Destination $script:Source
        $script:Installer = Join-Path $script:CaseRoot 'retained-install.ps1'
        Copy-Item -LiteralPath (Join-Path $script:Repository 'scripts/install.ps1') -Destination $script:Installer
        $script:SavedEnvironment = @{}
        $values = @{
            XDG_CONFIG_HOME=(Join-Path $script:CaseRoot 'xdg-config'); XDG_DATA_HOME=(Join-Path $script:CaseRoot 'xdg-data')
            XDG_CACHE_HOME=(Join-Path $script:CaseRoot 'xdg-cache'); XDG_STATE_HOME=(Join-Path $script:CaseRoot 'xdg-state')
            OPENCODE_TEST_HOME=(Join-Path $script:CaseRoot 'home'); OPENCODE_CONFIG=(Join-Path $script:CaseRoot 'empty.json')
            OPENCODE_CONFIG_CONTENT='{}'; OPENCODE_CONFIG_DIR=$null; OPENCODE_DISABLE_AUTOUPDATE='1'; OPENCODE_DISABLE_MODELS_FETCH='1'
        }
        foreach ($key in $values.Keys) { $script:SavedEnvironment[$key] = [Environment]::GetEnvironmentVariable($key, 'Process'); [Environment]::SetEnvironmentVariable($key, $values[$key], 'Process') }
        foreach ($key in @('XDG_CONFIG_HOME', 'XDG_DATA_HOME', 'XDG_CACHE_HOME', 'XDG_STATE_HOME', 'OPENCODE_TEST_HOME')) { [void][IO.Directory]::CreateDirectory($values[$key]) }
        Write-DiscoveryText $env:OPENCODE_CONFIG "{}`n"
        $version = Invoke-CIProcess $env:MAIL_LENS_OPENCODE @('--version') $script:Project
        if ($version.Code -ne 0 -or $version.Out.Trim() -cne '1.18.30') { throw 'Unexpected OpenCode version.' }
        Write-Host "Native OpenCode version: $($version.Out.Trim())."
    }
    AfterEach {
        foreach ($key in $script:SavedEnvironment.Keys) { [Environment]::SetEnvironmentVariable($key, $script:SavedEnvironment[$key], 'Process') }
        if ([IO.Directory]::Exists($script:CaseRoot)) { [IO.Directory]::Delete($script:CaseRoot, $true) }
    }
    It 'discovers install/status/update/rollback/relocation/removal in <Scope>' -ForEach @(@{Scope='default'}, @{Scope='User'}, @{Scope='Project'}) {
        Select-DiscoveryScope $Scope
        (Invoke-Installer).Code | Should -Be 0
        Assert-DiscoveredPackage
        $before = Get-DiscoverySnapshot $script:Target
        foreach ($action in @('Status', 'Install')) { (Invoke-Installer $action).Code | Should -Be 0; (Get-DiscoverySnapshot $script:Target) | Should -BeExactly $before; Assert-DiscoveredPackage }
        $old = Join-Path $script:CaseRoot 'old source'; Copy-Item -LiteralPath $script:Source -Destination $old -Recurse
        $entry = Join-Path $script:Source 'skills/mail-lens/SKILL.md'
        Write-DiscoveryText $entry ([IO.File]::ReadAllText($entry, $script:Utf8) + "`nSynthetic updated content.`n")
        Write-DiscoveryText (Join-Path $script:Source 'skills/mail-lens/addition.txt') 'Synthetic resource'
        (Invoke-Installer 'Update').Code | Should -Be 0; Assert-DiscoveredPackage
        (Invoke-Installer 'Update' $old).Code | Should -Be 0; Assert-DiscoveredPackage $script:Target $old
        [IO.File]::Exists((Join-Path $script:Target 'addition.txt')) | Should -BeFalse
        $moved = Join-Path $script:CaseRoot 'moved source'; [IO.Directory]::Move($old, $moved)
        Assert-DiscoveredPackage $script:Target $moved
        [IO.Directory]::Delete($moved, $true); [IO.Directory]::Delete($script:Source, $true)
        $before = Get-DiscoverySnapshot $script:Target
        (Invoke-Installer 'Status').Code | Should -Be 0; (Get-DiscoverySnapshot $script:Target) | Should -BeExactly $before
        @(Get-DiscoveredSkills | Where-Object name -CEQ 'mail-lens').Count | Should -Be 1
        (Invoke-Installer 'Uninstall').Code | Should -Be 0
        [IO.Directory]::Exists($script:Target) | Should -BeFalse
        @(Get-DiscoveredSkills | Where-Object name -CEQ 'mail-lens').Count | Should -Be 0
    }
    It 'keeps a valid skill discoverable after <Receipt> receipt rejects update/uninstall' -ForEach @(@{Receipt='missing'}, @{Receipt='corrupt'}) {
        Select-DiscoveryScope 'Project'; (Invoke-Installer).Code | Should -Be 0
        $path = Join-Path $script:Target '.mail-lens-install.json'
        if ($Receipt -eq 'missing') { [IO.File]::Delete($path) } else { Write-DiscoveryText $path 'not JSON' }
        $before = Get-DiscoverySnapshot $script:Target
        foreach ($action in @('Update', 'Uninstall')) { (Invoke-Installer $action).Code | Should -Be 1; (Get-DiscoverySnapshot $script:Target) | Should -BeExactly $before; Assert-DiscoveredPackage }
    }
    It 'removes only selected duplicate scope and preserves other skills/configuration' {
        Select-DiscoveryScope 'default'; (Invoke-Installer).Code | Should -Be 0
        $globalTarget = $script:Target
        $globalBefore = Get-DiscoverySnapshot $globalTarget
        Select-DiscoveryScope 'Project'; (Invoke-Installer).Code | Should -Be 0
        $other = Join-Path $script:Project '.opencode/skills/synthetic-other'
        [void][IO.Directory]::CreateDirectory($other)
        Write-DiscoveryText (Join-Path $other 'SKILL.md') "---`nname: synthetic-other`ndescription: Synthetic sentinel`n---`nSynthetic preserved skill.`n"
        $config = Join-Path $script:Project '.opencode/opencode.json'; Write-DiscoveryText $config "{}`n"
        $otherBefore = Get-DiscoverySnapshot $other
        $winner = @(Get-DiscoveredSkills | Where-Object name -CEQ 'mail-lens')
        $winner.Count | Should -Be 1
        # OpenCode may normalize its synthetic config; snapshot after discovery.
        $configBefore = [IO.File]::ReadAllBytes($config)
        @((Join-Path $script:Target 'SKILL.md'), (Join-Path $globalTarget 'SKILL.md')) | Should -Contain ([IO.Path]::GetFullPath($winner[0].location))
        (Invoke-Installer 'Uninstall').Code | Should -Be 0
        [IO.Directory]::Exists($script:Target) | Should -BeFalse
        (Get-DiscoverySnapshot $globalTarget) | Should -BeExactly $globalBefore
        Assert-DiscoveredPackage $globalTarget
        (Get-DiscoverySnapshot $other) | Should -BeExactly $otherBefore
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($config)) | Should -BeExactly ([Convert]::ToBase64String($configBefore))
        @(Get-DiscoveredSkills | Where-Object name -CEQ 'synthetic-other').Count | Should -Be 1
        Select-DiscoveryScope 'default'; (Invoke-Installer 'Uninstall').Code | Should -Be 0
        @(Get-DiscoveredSkills | Where-Object name -CEQ 'mail-lens').Count | Should -Be 0
    }
}
