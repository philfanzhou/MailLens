BeforeAll {
    $script:Repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    . (Join-Path $script:Repository 'scripts/check-package.ps1')
    . (Join-Path $script:Repository 'scripts/ci-process.ps1')
    $script:Utf8 = New-Object Text.UTF8Encoding($false, $true)
    $script:PowerShell = (Get-Command powershell.exe -CommandType Application -ErrorAction Stop).Source
    function Write-CIText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:Utf8) }
    function Invoke-NegativeTests([string]$Path, [string]$Manifest = $env:MAIL_LENS_PESTER_MANIFEST, [string]$Required = '') {
        Invoke-CIProcess $script:PowerShell @('-NoProfile', '-STA', '-File', (Join-Path $script:Repository 'scripts/invoke-tests.ps1'), '-TestPath', $Path, '-PesterManifest', $Manifest, '-RequiredContainers', $Required) $script:CaseRoot
    }
}
Describe 'CI failure propagation' {
    BeforeEach {
        $script:CaseRoot = Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-ci-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($script:CaseRoot)
        $script:Copy = Join-Path $script:CaseRoot 'source'
        [void][IO.Directory]::CreateDirectory($script:Copy)
        foreach ($name in @('skills', 'scripts', 'VERSION', 'LICENSE')) { Copy-Item -LiteralPath (Join-Path $script:Repository $name) -Destination $script:Copy -Recurse }
    }
    AfterEach { if ([IO.Directory]::Exists($script:CaseRoot)) { [IO.Directory]::Delete($script:CaseRoot, $true) } }
    It 'accepts the actual package' { { Test-MailLensPackage $script:Copy } | Should -Not -Throw }
    It 'rejects broken package <Fault>' -ForEach @(
        @{Fault='missing-entry'}, @{Fault='missing-frontmatter'}, @{Fault='empty-description'},
        @{Fault='duplicate-field'}, @{Fault='wrong-name'}, @{Fault='bad-field'}, @{Fault='missing-resource'},
        @{Fault='escaping-resource'}, @{Fault='missing-version'}, @{Fault='bad-version'}, @{Fault='duplicate-version'},
        @{Fault='missing-bom'}, @{Fault='malformed-utf8'}, @{Fault='parse-error'}, @{Fault='whitespace'}, @{Fault='comment-language'}) {
        $entry = Join-Path $script:Copy 'skills/mail-lens/SKILL.md'
        $text = [IO.File]::ReadAllText($entry, $script:Utf8)
        switch ($Fault) {
            'missing-entry' { [IO.File]::Delete($entry) }
            'missing-frontmatter' { Write-CIText $entry "# Synthetic skill`n" }
            'empty-description' { Write-CIText $entry "---`nname: mail-lens`ndescription: `n---`n" }
            'duplicate-field' { Write-CIText $entry "---`nname: mail-lens`nname: mail-lens`ndescription: synthetic`n---`n" }
            'wrong-name' { Write-CIText $entry "---`nname: wrong`ndescription: synthetic`n---`n" }
            'bad-field' { Write-CIText $entry "---`nname: mail-lens`ndescription: [array]`n---`n" }
            'missing-resource' { [IO.File]::Delete((Join-Path $script:Copy 'skills/mail-lens/references/workflow.md')) }
            'escaping-resource' { Write-CIText $entry ($text + "`n[escape](../../VERSION)`n") }
            'missing-version' { [IO.File]::Delete((Join-Path $script:Copy 'VERSION')) }
            'bad-version' { Write-CIText (Join-Path $script:Copy 'VERSION') "latest`n" }
            'duplicate-version' { Write-CIText $entry ($text.Replace('name: mail-lens', "name: mail-lens`nversion: 0.1.0")) }
            'missing-bom' { Write-CIText (Join-Path $script:Copy 'scripts/install.ps1') "# synthetic`n" }
            'malformed-utf8' { [IO.File]::WriteAllBytes($entry, [byte[]]@(255,254,255)) }
            'parse-error' { [IO.File]::WriteAllText((Join-Path $script:Copy 'scripts/install.ps1'), "function {`n", (New-Object Text.UTF8Encoding($true, $true))) }
            'whitespace' { Write-CIText $entry ($text + "bad `n") }
            'comment-language' { [IO.File]::WriteAllText((Join-Path $script:Copy 'scripts/install.ps1'), ('# ' + [char]0x4e2d + "`n"), (New-Object Text.UTF8Encoding($true, $true))) }
        }
        { Test-MailLensPackage $script:Copy } | Should -Throw
    }
    It 'rejects missing command/path, empty containers and required-container deletion' -ForEach @(@{Fault='command'}, @{Fault='path'}, @{Fault='empty'}, @{Fault='required'}) {
        $testPath = Join-Path $script:CaseRoot 'tests'
        [void][IO.Directory]::CreateDirectory($testPath)
        $manifest = $env:MAIL_LENS_PESTER_MANIFEST
        $required = ''
        switch ($Fault) {
            'command' { $manifest = Join-Path $script:CaseRoot 'missing.psd1' }
            'path' { $testPath = Join-Path $script:CaseRoot 'missing' }
            'empty' { Write-CIText (Join-Path $testPath 'Empty.Tests.ps1') "Describe 'empty' { }`n" }
            'required' { Write-CIText (Join-Path $testPath 'Pass.Tests.ps1') "Describe 'pass' { It 'works' { 1 | Should -Be 1 } }`n"; $required = 'Missing.Tests.ps1' }
        }
        (Invoke-NegativeTests $testPath $manifest $required).Code | Should -Be 1
    }
    It 'propagates deliberately failing and skipped synthetic tests' -ForEach @(@{Kind='failure'}, @{Kind='skip'}) {
        $testPath = Join-Path $script:CaseRoot 'tests'; [void][IO.Directory]::CreateDirectory($testPath)
        if ($Kind -eq 'failure') { $test = "Describe 'negative' { It 'fails' { 1 | Should -Be 2 } }`n" }
        else { $test = "Describe 'negative' { It 'skips' -Skip { 1 | Should -Be 1 } }`n" }
        Write-CIText (Join-Path $testPath 'Negative.Tests.ps1') $test
        (Invoke-NegativeTests $testPath).Code | Should -Be 1
    }
    It 'rejects absent CLI executable and times out only its owned child' {
        { Invoke-CIProcess (Join-Path $script:CaseRoot 'absent.exe') @() $script:CaseRoot } | Should -Throw '*missing*'
        $scriptPath = Join-Path $script:CaseRoot 'sleep.ps1'; Write-CIText $scriptPath "Start-Sleep -Seconds 30`n"
        { Invoke-CIProcess $script:PowerShell @('-NoProfile', '-File', $scriptPath) $script:CaseRoot 1 } | Should -Throw '*timed out*'
    }
    It 'rejects non-JSON, malformed records, missing target and wrong target path/content' {
        { ConvertFrom-CISkills 'not json' } | Should -Throw
        { ConvertFrom-CISkills '[{"name":"mail-lens"}]' } | Should -Throw
        $path = Join-Path $script:CaseRoot 'SKILL.md'
        { Assert-CITarget @() $path 'body' 'description' } | Should -Throw
        $record = [pscustomobject]@{name='mail-lens';location=(Join-Path $script:CaseRoot 'wrong.md');content='body';description='description'}
        { Assert-CITarget @($record) $path 'body' 'description' } | Should -Throw '*path*'
        $record.location = $path; $record.content = 'wrong'
        { Assert-CITarget @($record) $path 'body' 'description' } | Should -Throw '*content*'
    }
}
