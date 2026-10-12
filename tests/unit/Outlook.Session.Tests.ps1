BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/MailLens.Outlook.psm1'
    Import-Module $modulePath -Force
    . (Join-Path $PSScriptRoot '../fixtures/com/SessionDouble.ps1')
    $native = @{ isWindows = $true; powerShellVersion = '5.1.26100.33438'; edition = 'Desktop'; apartmentState = 'STA'; timeZone = 'Synthetic Standard Time' }
    function Request($extra = '') { ConvertFrom-Json ('{"schemaVersion":1,"operation":"folders"' + $extra + '}') }
}

Describe 'Read-only session and folder catalog' {
    It 'discovers all mounted stores and nested, custom, deleted and junk folders in deterministic order' {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be complete
        $result.data.stores.Count | Should -Be 2
        ($result.data.folders.folderId -join ',') | Should -Be 'f-root,f-projects,f-nested,f-calendar,f-calendar-mail,f-deleted,f-junk,f-archive'
        $result.data.coverage.complete | Should -BeTrue
        $result.data.folders.Count | Should -Be 8
        Assert-AllReleased $double
        @($double.state.calls | Where-Object { $_ -match '\.(Items|Body|Quit|Logoff|Logon|Display|Save|Delete|Send|Move)(\.|$)' }).Count | Should -Be 0
    }
    It 'represents the <Mode> session path without profile switching or termination' -ForEach @(@{ Mode = 'existing' }, @{ Mode = 'started' }) {
        $double = New-SessionDouble -Mode $Mode
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.data.sessionMode | Should -BeExactly $Mode
        $double.state.connections | Should -Be 1
        Assert-AllReleased $double
    }
    It 'traverses non-mail parents and does not inspect their mixed items' {
        $double = New-SessionDouble
        # Any Items access is forbidden by the double, irrespective of mixed contents.
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        @($result.data.folders | Where-Object { $_.defaultItemType -eq 1 }).Count | Should -Be 1
        $result.data.folders.folderId | Should -Contain 'f-calendar-mail'
        Assert-AllReleased $double
    }
    It 'handles no mounted stores as a complete empty catalog' {
        $double = New-SessionDouble -Empty
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be complete
        $result.data.stores.Count | Should -Be 0
        $result.data.folders.Count | Should -Be 0
        Assert-AllReleased $double
    }
    It 'retains IDs and Unicode through JSON without serializing any adapter reference' {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $json = ConvertTo-Json -InputObject $result -Depth 30
        $copy = ConvertFrom-Json $json
        $copy.data.stores[0].storeName | Should -BeExactly '工作📮'
        $copy.data.folders[2].folderPath | Should -BeExactly '\\工作📮\Projects\合同'
        $copy.data.folders[2].storeId | Should -BeExactly 's-a'
        $json | Should -Not -Match 'refId|properties|methods|System.__ComObject'
        Assert-AllReleased $double
    }
    It 'breaks duplicate/cyclic folder visits using store plus folder identity' {
        $double = New-SessionDouble
        $double.state.root.properties.Folders.methods.Item += $double.state.root
        $double.state.root.properties.Folders.properties.Count++
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be complete
        $result.data.folders.Count | Should -Be 8
        Assert-AllReleased $double
    }
    It 'diagnoses using a serializable native/runtime/session envelope' {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (ConvertFrom-Json '{"schemaVersion":1,"operation":"diagnose"}') $double.adapter $native
        $result.status | Should -Be complete
        $result.data.runtime.apartmentState | Should -BeExactly 'STA'
        $result.startedAt | Should -Not -BeNullOrEmpty
        $result.endedAt | Should -Not -BeNullOrEmpty
        Assert-AllReleased $double
    }
}

Describe 'Exact scope resolution' {
    It 'selects an exact mailbox ID/name and excludes the other store' -ForEach @(@{ Selector = 's-a' }, @{ Selector = '工作📮' }) {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request (',"mailboxes":["' + $Selector + '"]')) $double.adapter $native
        $result.status | Should -Be complete
        $result.scope.stores.Count | Should -Be 1
        $result.scope.folders.Count | Should -Be 7
        Assert-AllReleased $double
    }
    It 'returns actionable ambiguity and catalog choices for duplicate store display names' {
        $double = New-SessionDouble -DuplicateNames
        $result = Invoke-MailLensRequest (Request ',"mailboxes":["工作📮"]') $double.adapter $native
        $result.status | Should -Be failed
        $result.errors[0].code | Should -BeExactly 'ambiguous_selection'
        $result.data.stores.storeId | Should -Contain 's-b'
        $result.scope.stores.Count | Should -Be 0
        Assert-AllReleased $double
    }
    It 'never broadens an unknown mailbox/folder or a folder outside selected stores' -ForEach @(
        @{ Extra = ',"mailboxes":["missing"]' },
        @{ Extra = ',"folders":["missing"]' },
        @{ Extra = ',"mailboxes":["s-b"],"folders":[{"storeId":"s-a","folderId":"f-root"}]' }
    ) {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request $Extra) $double.adapter $native
        $result.status | Should -Be failed
        $result.errors[0].code | Should -BeExactly 'unknown_selection'
        $result.scope.folders.Count | Should -Be 0
        Assert-AllReleased $double
    }
    It 'selects a folder identity with explicit <Descendants> descendants' -ForEach @(@{ Descendants = 'true'; Count = 2 }, @{ Descendants = 'false'; Count = 1 }) {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request (',"folders":[{"storeId":"s-a","folderId":"f-projects"}],"includeDescendants":' + $Descendants)) $double.adapter $native
        $result.status | Should -Be complete
        $result.scope.folders.Count | Should -Be $Count
        $result.scope.stores.Count | Should -Be 1
        Assert-AllReleased $double
    }
    It 'selects an exact Unicode folder path with descendants by default' {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request ',"folders":["\\\\工作📮\\Projects"]') $double.adapter $native
        $result.status | Should -Be complete
        $result.scope.folders.Count | Should -Be 2
        Assert-AllReleased $double
    }
    It 'rejects ambiguous exact paths while preserving the identity choices' {
        $double = New-SessionDouble
        $double.state.archiveRoot.properties.FolderPath = '\\工作📮\Projects'
        $result = Invoke-MailLensRequest (Request ',"folders":["\\\\工作📮\\Projects"]') $double.adapter $native
        $result.errors[0].code | Should -BeExactly 'ambiguous_selection'
        Assert-AllReleased $double
    }
}

Describe 'Failure, cancellation and reference ownership' {
    It 'continues other folders/stores after denial at <Failure>' -ForEach @(
        @{ Failure = 'store-a.GetRootFolder' }, @{ Failure = 'f-projects.Folders' },
        @{ Failure = 'f-root-children.Item.1' }, @{ Failure = 'f-projects.DefaultItemType' },
        @{ Failure = 'f-projects.EntryID' }, @{ Failure = 'f-projects.FolderPath' },
        @{ Failure = 'store-a.DisplayName' }
    ) {
        $double = New-SessionDouble -Failures @{ $Failure = $true }
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be partial
        $result.data.folders.folderId | Should -Contain 'f-archive'
        $result.warnings.Count | Should -BeGreaterThan 0
        $result.data.coverage.complete | Should -BeFalse
        Assert-AllReleased $double
    }
    It 'fails fully denied discovery rather than returning an empty success' {
        $double = New-SessionDouble -Failures @{ 'store-a.GetRootFolder' = $true; 'store-b.GetRootFolder' = $true }
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be failed
        $result.warnings.Count | Should -Be 2
        Assert-AllReleased $double
    }
    It 'fails when all root metadata is unreadable despite acquiring the roots' {
        $double = New-SessionDouble -Empty
        # Replace the mounted store collection with only one wholly denied root.
        $ordinary = New-SessionDouble -Failures @{ 'f-root.FolderPath' = $true; 'store-b.GetRootFolder' = $true }
        $ordinary.state.root.properties.Folders.methods.Item = @()
        $ordinary.state.root.properties.Folders.properties.Count = 0
        $result = Invoke-MailLensRequest (Request) $ordinary.adapter $native
        $result.status | Should -Be failed
        $result.data.coverage.readableFolders | Should -Be 0
        Assert-AllReleased $ordinary
    }
    It 'reports an explicitly selected denied store as failed with its gap' {
        $double = New-SessionDouble -Failures @{ 'store-a.GetRootFolder' = $true }
        $result = Invoke-MailLensRequest (Request ',"mailboxes":["s-a"]') $double.adapter $native
        $result.status | Should -Be failed
        $result.scope.stores[0].readable | Should -BeFalse
        Assert-AllReleased $double
    }
    It 'releases application/session/collections on <Failure> failure' -ForEach @(@{ Failure = 'application.GetNamespace' }, @{ Failure = 'session.Stores' }, @{ Failure = 'stores.Count' }) {
        $double = New-SessionDouble -Failures @{ $Failure = $true }
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be failed
        Assert-AllReleased $double
    }
    It 'reports missing Classic COM or denied profile prompts without leaking raw exceptions' {
        $double = New-SessionDouble -Failures @{ Connect = $true }
        $result = Invoke-MailLensRequest (Request) $double.adapter $native
        $result.status | Should -Be failed
        $result.errors[0].code | Should -BeExactly 'outlook_unavailable'
        $result.errors[0].message | Should -Not -Match 'Synthetic'
        Assert-AllReleased $double
    }
    It 'cancels before COM with exit 130' {
        $double = New-SessionDouble
        $result = Invoke-MailLensRequest (Request) $double.adapter $native { $true }
        $result.status | Should -Be cancelled
        (Get-MailLensExitCode $result.status) | Should -Be 130
        $double.state.connections | Should -Be 0
    }
    It 'cancels between child visits, preserves partial coverage and releases all acquired references' {
        $double = New-SessionDouble
        $state = $double.state
        $cancel = { $state.calls.Contains('f-projects-children.Release') }.GetNewClosure()
        $result = Invoke-MailLensRequest (Request) $double.adapter $native $cancel
        $result.status | Should -Be cancelled
        $result.data.folders.Count | Should -BeGreaterThan 0
        $result.data.folders.Count | Should -BeLessThan 8
        $result.data.coverage.complete | Should -BeFalse
        Assert-AllReleased $double
    }
    It 'uses stable documented exit codes' {
        (Get-MailLensExitCode complete) | Should -Be 0
        (Get-MailLensExitCode partial) | Should -Be 2
        (Get-MailLensExitCode failed) | Should -Be 1
    }
}

Describe 'Validation before Outlook' {
    It 'rejects invalid request <Json> without connecting' -ForEach @(
        @{ Json = 'null' }, @{ Json = '[]' }, @{ Json = '{"schemaVersion":2,"operation":"folders"}' },
        @{ Json = '{"schemaVersion":"1","operation":"folders"}' }, @{ Json = '{"schemaVersion":1,"operation":"Folders"}' },
        @{ Json = '{"schemaVersion":1,"operation":"unknown"}' }, @{ Json = '{"schemaVersion":1,"operation":"search"}' },
        @{ Json = '{"schemaVersion":1,"operation":"get"}' }, @{ Json = '{"schemaVersion":1,"operation":"diagnose","mailboxes":["s-a"]}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","unknown":true}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","mailboxes":[]}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","mailboxes":"s-a"}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","mailboxes":[null]}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","folders":[{"folderId":"f-root"}]}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","folders":[{"storeId":"s-a","folderId":"f-root","path":"x"}]}' },
        @{ Json = '{"schemaVersion":1,"operation":"folders","includeDescendants":"false"}' }
    ) {
        $double = New-SessionDouble
        $request = ConvertFrom-Json $Json
        $result = Invoke-MailLensRequest $request $double.adapter $native
        $result.status | Should -Be failed
        $double.state.connections | Should -Be 0
        Assert-AllReleased $double
    }
    It 'guards <Field> = <Value> before COM' -ForEach @(
        @{ Field = 'isWindows'; Value = $false; Code = 'unsupported_platform' },
        @{ Field = 'edition'; Value = 'Core'; Code = 'unsupported_powershell' },
        @{ Field = 'powerShellVersion'; Value = '7.5.0'; Code = 'unsupported_powershell' },
        @{ Field = 'apartmentState'; Value = 'MTA'; Code = 'unsupported_apartment' }
    ) {
        $double = New-SessionDouble
        $runtime = $native.Clone(); $runtime[$Field] = $Value
        $result = Invoke-MailLensRequest (Request) $double.adapter $runtime
        $result.errors[0].code | Should -BeExactly $Code
        $double.state.connections | Should -Be 0
    }
}

Describe 'Atomic UTF-8 request/result files' {
    BeforeEach {
        $directory = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($directory)
        $requestPath = Join-Path $directory 'request.json'
        $outputPath = Join-Path $directory 'result.json'
        [IO.File]::WriteAllText($requestPath, '{"schemaVersion":1,"operation":"folders"}', (New-Object Text.UTF8Encoding($false)))
        $double = New-SessionDouble
    }
    It 'writes a parseable Unicode response, no BOM and no leftover temp/lock' {
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 0
        $bytes = [IO.File]::ReadAllBytes($outputPath)
        $bytes[0] | Should -Be 123
        $json = [IO.File]::ReadAllText($outputPath, [Text.Encoding]::UTF8)
        (ConvertFrom-Json $json).data.stores[0].storeName | Should -BeExactly '工作📮'
        @(Get-ChildItem $directory -Force).Count | Should -Be 2
        Assert-AllReleased $double
    }
    It 'produces a failed English envelope for malformed JSON before COM' {
        [IO.File]::WriteAllText($requestPath, '{"bad":', [Text.Encoding]::UTF8)
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 1
        $result = ConvertFrom-Json ([IO.File]::ReadAllText($outputPath))
        $result.errors[0].code | Should -BeExactly 'invalid_request_file'
        $double.state.connections | Should -Be 0
    }
    It 'accepts an optional UTF-8 request BOM' {
        [IO.File]::WriteAllText($requestPath, '{"schemaVersion":1,"operation":"folders"}', (New-Object Text.UTF8Encoding($true)))
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 0
        Assert-AllReleased $double
    }
    It 'rejects malformed UTF-8 and UTF-16 input before COM' -ForEach @(@{ Encoding = 'malformed' }, @{ Encoding = 'utf16' }) {
        if ($Encoding -eq 'malformed') { [IO.File]::WriteAllBytes($requestPath, [byte[]]@(123, 0xFF, 125)) }
        else { [IO.File]::WriteAllText($requestPath, '{"schemaVersion":1,"operation":"folders"}', [Text.Encoding]::Unicode) }
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 1
        $double.state.connections | Should -Be 0
    }
    It 'produces a failed envelope for an unreadable/missing request before COM' {
        [IO.File]::Delete($requestPath)
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 1
        $double.state.connections | Should -Be 0
    }
    It 'refuses existing output without altering its bytes or connecting' {
        [IO.File]::WriteAllText($outputPath, 'original')
        { Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native } | Should -Throw
        [IO.File]::ReadAllText($outputPath) | Should -BeExactly 'original'
        $double.state.connections | Should -Be 0
    }
    It 'refuses a reused destination while another invocation holds its reservation' {
        $lock = [IO.File]::Open(($outputPath + '.maillens-lock'), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            { Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native } | Should -Throw
            $double.state.connections | Should -Be 0
        } finally { $lock.Dispose(); [IO.File]::Delete($outputPath + '.maillens-lock') }
    }
    It 'refuses relative and blocked output paths before COM' {
        { Invoke-MailLensFileRequest $requestPath 'relative.json' $double.adapter $native } | Should -Throw
        { Invoke-MailLensFileRequest $requestPath 'C:relative.json' $double.adapter $native } | Should -Throw
        { Invoke-MailLensFileRequest $requestPath (Join-Path $directory 'missing/result.json') $double.adapter $native } | Should -Throw
        $double.state.connections | Should -Be 0
    }
    It 'does not overwrite a file created by another writer after validation' {
        $state = $double.state
        $connect = $double.adapter.Connect
        $destination = $outputPath
        $double.adapter.Connect = {
            [IO.File]::WriteAllText($destination, 'other-writer')
            & $connect
        }.GetNewClosure()
        { Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native } | Should -Throw
        [IO.File]::ReadAllText($outputPath) | Should -BeExactly 'other-writer'
        @(Get-ChildItem $directory -Filter '.maillens-*' -Force).Count | Should -Be 0
        Assert-AllReleased $double
    }
    It 'maps partial and cancelled responses into result files and exit codes' {
        $double = New-SessionDouble -Failures @{ 'f-projects.Folders' = $true }
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native) | Should -Be 2
        [IO.File]::Delete($outputPath)
        (Invoke-MailLensFileRequest $requestPath $outputPath $double.adapter $native { $true }) | Should -Be 130
        (ConvertFrom-Json ([IO.File]::ReadAllText($outputPath))).status | Should -BeExactly 'cancelled'
        Assert-AllReleased $double
    }
    It 'launches the shipped entrypoint in MTA and returns a failed envelope without COM' {
        $entrypoint = Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/outlook.ps1'
        $printed = @(& (Join-Path $PSHOME 'powershell.exe') -NoProfile -MTA -File $entrypoint -RequestPath $requestPath -OutputPath $outputPath 2>&1)
        $LASTEXITCODE | Should -Be 1
        $result = ConvertFrom-Json ([IO.File]::ReadAllText($outputPath))
        $result.errors[0].code | Should -BeExactly 'unsupported_apartment'
        $printed.Count | Should -Be 0
    }
    It 'allows concurrent independent CLI invocations with distinct result destinations' {
        [IO.File]::WriteAllText($requestPath, '{bad JSON', [Text.Encoding]::UTF8)
        $entrypoint = Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/outlook.ps1'
        $otherOutput = Join-Path $directory 'other-result.json'
        $exe = Join-Path $PSHOME 'powershell.exe'
        $firstArgs = @('-NoProfile', '-STA', '-File', ('"' + $entrypoint + '"'), '-RequestPath', ('"' + $requestPath + '"'), '-OutputPath', ('"' + $outputPath + '"'))
        $secondArgs = @('-NoProfile', '-STA', '-File', ('"' + $entrypoint + '"'), '-RequestPath', ('"' + $requestPath + '"'), '-OutputPath', ('"' + $otherOutput + '"'))
        $first = Start-Process $exe -ArgumentList $firstArgs -PassThru -NoNewWindow
        $second = Start-Process $exe -ArgumentList $secondArgs -PassThru -NoNewWindow
        try {
            $first.WaitForExit(30000) | Should -BeTrue
            $second.WaitForExit(30000) | Should -BeTrue
            $first.ExitCode | Should -Be 1
            $second.ExitCode | Should -Be 1
            foreach ($path in @($outputPath, $otherOutput)) {
                (ConvertFrom-Json ([IO.File]::ReadAllText($path))).errors[0].code | Should -BeExactly 'invalid_request_file'
            }
            @(Get-ChildItem $directory -Force).Count | Should -Be 3
        } finally { $first.Dispose(); $second.Dispose() }
    }
}
