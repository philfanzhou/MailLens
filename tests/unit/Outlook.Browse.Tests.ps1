BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/MailLens.Outlook.psm1') -Force
    . (Join-Path $PSScriptRoot '../fixtures/com/SessionDouble.ps1')
    . (Join-Path $PSScriptRoot '../fixtures/com/BrowseDouble.ps1')
    $native = @{ isWindows = $true; powerShellVersion = '5.1.26100.33438'; edition = 'Desktop'; apartmentState = 'STA'; timeZone = 'UTC' }
    function BrowseRequest($Extra = '') {
        $request = ConvertFrom-Json ('{"schemaVersion":2,"operation":"browse"' + $Extra + '}')
        # PowerShell 7's JSON integers are Int64; emulate the native 5.1 parser for the private seam.
        foreach ($p in @($request.PSObject.Properties)) { if ($p.Value -is [long] -and $p.Value -le [int]::MaxValue -and $p.Value -ge [int]::MinValue) { $p.Value = [int]$p.Value } }
        return $request
    }
    function Browse($Double, $Request = (BrowseRequest), $Cancel = { $false }, $Clock = { [double][Diagnostics.Stopwatch]::GetTimestamp() / [Diagnostics.Stopwatch]::Frequency }, $Progress = {}) {
        Invoke-MailLensRequest -Request $Request -Adapter $Double.adapter -Runtime $native -CancellationRequested $Cancel -Clock $Clock -Progress $Progress
    }
    function Assert-NoBroadOrBodyReads($Double) {
        @($Double.state.calls | Where-Object { $_ -match '\.(Stores|GetRootFolder|Folders|Body|LastModificationTime|Sent|SenderName|SenderEmailAddress|Recipients|Attachments|Quit|Logon|Logoff|Display|Save|Delete|Move|Send)(\.|$)' }).Count | Should -Be 0
    }
}

Describe 'Direct limited Inbox browsing' {
    It 'fetches only ten of 1000 readable messages in descending ReceivedTime order' {
        $double = New-BrowseDouble
        $result = Browse $double
        $result.status | Should -Be complete
        $result.data.browse.returnedCount | Should -Be 10
        $result.data.browse.visitedCount | Should -Be 10
        $result.data.browse.limit | Should -Be 10
        $result.data.browse.maxItems | Should -Be 200
        $result.data.browse.maxDurationSeconds | Should -Be 30
        $result.data.browse.ordering | Should -Be receivedTimeDescending
        $result.data.browse.summaries[0].entryId | Should -BeExactly 'browse-1000'
        $result.data.browse.summaries[9].entryId | Should -BeExactly 'browse-991'
        $result.data.browse.summaries[0].subject | Should -BeExactly '合成📮 1000'
        $result.data.browse.summaries[0].storeId | Should -BeExactly 's-a'
        $result.data.browse.summaries[0].folderId | Should -BeExactly 'inbox'
        $result.data.browse.summaries[0].timestampField | Should -BeExactly 'ReceivedTime'
        $result.data.browse.summaries[0].preview | Should -BeNullOrEmpty
        $result.data.browse.summaries[0].previewSource | Should -BeNullOrEmpty
        $result.data.browse.scanComplete | Should -BeFalse
        $result.data.browse.moreAvailable | Should -BeTrue
        $result.data.browse.stopReason | Should -Be limit
        @($double.state.calls | Where-Object { $_ -match '^inbox-items.Item\.' }).Count | Should -Be 10
        @($double.state.calls | Where-Object { $_ -eq 'store-a.GetDefaultFolder' }).Count | Should -Be 1
        Assert-NoBroadOrBodyReads $double
        Assert-AllReleased $double
    }
    It 'reaches ordinary end for <Count> messages' -ForEach @(@{ Count = 0 }, @{ Count = 3 }, @{ Count = 10 }) {
        $double = New-BrowseDouble -Count $Count
        $result = Browse $double
        $result.status | Should -Be complete
        $result.data.browse.returnedCount | Should -Be $Count
        $result.data.browse.scanComplete | Should -BeTrue
        $result.data.browse.moreAvailable | Should -BeFalse
        $result.data.browse.stopReason | Should -Be end
        Assert-AllReleased $double
    }
    It 'uses exact acquisition only and ignores unrelated store/root failures' {
        $double = New-BrowseDouble -Failures @{ 'session.Stores' = $true; 'store-b.GetRootFolder' = $true; 'f-root.Folders' = $true; 'session.DefaultStore' = $true }
        $result = Browse $double (BrowseRequest ',"storeId":"s-a","folderId":"inbox"')
        $result.status | Should -Be complete
        $double.state.calls | Should -Contain 'session.GetFolderFromID'
        $double.state.calls | Should -Not -Contain 'session.DefaultStore'
        $double.state.calls | Should -Not -Contain 'store-a.GetDefaultFolder'
        Assert-NoBroadOrBodyReads $double
        Assert-AllReleased $double
    }
    It 'does not let an unrelated denied store/root/child affect default browse' {
        $double = New-BrowseDouble -Failures @{ 'session.Stores' = $true; 'store-b.GetRootFolder' = $true; 'f-root.Folders' = $true }
        (Browse $double).status | Should -Be complete
        Assert-NoBroadOrBodyReads $double
        Assert-AllReleased $double
    }
    It 'discloses sort denial as a limited sample without fallback enumeration' {
        $double = New-BrowseDouble -Failures @{ 'inbox-items.Sort' = $true }
        $result = Browse $double
        $result.status | Should -Be partial
        $result.data.browse.ordering | Should -Be collectionOrder
        $result.data.browse.summaries[0].entryId | Should -BeExactly 'browse-1'
        $result.warnings[0].code | Should -Be sort_unavailable
        @($double.state.calls | Where-Object { $_ -eq 'inbox-items.Sort' }).Count | Should -Be 1
        @($double.state.calls | Where-Object { $_ -match '^inbox-items.Item\.' }).Count | Should -Be 10
        Assert-AllReleased $double
    }
    It 'counts mixed classes and denied attempts against the item budget' {
        $double = New-BrowseDouble -Failures @{ 'inbox-items.Item.2' = $true }
        $double.state.messages[999].properties.Class = 26
        $result = Browse $double (BrowseRequest ',"limit":3,"maxItems":3')
        $result.status | Should -Be partial
        $result.data.browse.visitedCount | Should -Be 3
        $result.data.browse.returnedCount | Should -Be 1
        $result.data.browse.nonMailCount | Should -Be 1
        $result.data.browse.unreadableCount | Should -Be 1
        $result.data.browse.stopReason | Should -Be item_budget
        $result.data.browse.moreAvailable | Should -BeTrue
        $result.warningCount | Should -Be 2
        Assert-AllReleased $double
    }
    It 'retains useful identity/subject with a null optional date or display field at <Failure>' -ForEach @(
        @{ Failure = 'browse-1000.ReceivedTime'; Code = 'timestamp_unavailable' },
        @{ Failure = 'inbox.FolderPath'; Code = 'display_metadata_unavailable' },
        @{ Failure = 'store-a.DisplayName'; Code = 'display_metadata_unavailable' }
    ) {
        $double = New-BrowseDouble -Failures @{ $Failure = $true }
        $result = Browse $double
        $result.status | Should -Be partial
        $result.data.browse.returnedCount | Should -Be 10
        $result.data.browse.summaries[0].entryId | Should -Be browse-1000
        $result.warnings.code | Should -Contain $Code
        if ($Code -eq 'timestamp_unavailable') { $result.data.browse.summaries[0].timestamp | Should -BeNullOrEmpty; $result.data.browse.summaries[0].timestampField | Should -BeNullOrEmpty }
        Assert-AllReleased $double
    }
    It 'rejects an unusable item at <Failure> without substituting its identity' -ForEach @(
        @{ Failure = 'browse-1000.Class' }, @{ Failure = 'browse-1000.EntryID' }, @{ Failure = 'browse-1000.Subject' }, @{ Failure = 'browse-1000.Parent' }
    ) {
        $double = New-BrowseDouble -Failures @{ $Failure = $true }
        $result = Browse $double
        $result.status | Should -Be partial
        $result.data.browse.summaries.entryId | Should -Not -Contain browse-1000
        $result.data.browse.unreadableCount | Should -Be 1
        $result.data.browse.visitedCount | Should -Be 11
        Assert-AllReleased $double
    }
    It 'discards a changed exact item identity and never uses its same-subject identity' {
        $double = New-BrowseDouble
        $state = @{ reads = 0 }
        $double.state.actions['browse-1000.EntryID'] = { param($item) $state.reads++; if ($state.reads -eq 2) { $item.properties.EntryID = 'changed' } }.GetNewClosure()
        $result = Browse $double
        $result.status | Should -Be partial
        $result.data.browse.summaries.entryId | Should -Not -Contain changed
        $result.data.browse.summaries.entryId | Should -Not -Contain browse-1000
        Assert-AllReleased $double
    }
    It 'keeps exact warning totals with only two detailed samples per repeated gap' {
        $failures = @{}
        for ($i = 1; $i -le 200; $i++) { $failures['inbox-items.Item.' + $i] = $true }
        $double = New-BrowseDouble -Failures $failures
        $result = Browse $double
        $result.status | Should -Be partial
        $result.data.browse.visitedCount | Should -Be 200
        $result.data.browse.unreadableCount | Should -Be 200
        $result.warningCount | Should -Be 201
        @($result.warningSummary | Where-Object {$_.code -eq 'browse_item_unavailable'})[0].count | Should -Be 200
        $result.warnings.Count | Should -Be 3
        $result.warningSamplesTruncated | Should -BeTrue
        Assert-AllReleased $double
    }
}

Describe 'Browse acquisition failures do not widen scope' {
    It 'returns specific sanitized <Code> at <Failure>' -ForEach @(
        @{ Failure = 'session.DefaultStore'; Code = 'default_store_unavailable' },
        @{ Failure = 'store-a.GetDefaultFolder'; Code = 'default_inbox_unavailable' },
        @{ Failure = 'inbox.StoreID'; Code = 'folder_identity_unavailable' },
        @{ Failure = 'inbox.EntryID'; Code = 'folder_identity_unavailable' },
        @{ Failure = 'inbox.Items'; Code = 'items_unavailable' },
        @{ Failure = 'inbox-items.Count'; Code = 'items_unavailable' },
        @{ Failure = 'Connect'; Code = 'outlook_unavailable' },
        @{ Failure = 'application.GetNamespace'; Code = 'outlook_unavailable' }
    ) {
        $double = New-BrowseDouble -Failures @{ $Failure = $true }
        $result = Browse $double
        $result.status | Should -Be failed
        $result.errors[0].code | Should -Be $Code
        $result.errors[0].message | Should -Not -Match 'Synthetic|private|stack'
        Assert-NoBroadOrBodyReads $double
        Assert-AllReleased $double
    }
    It 'fails an absent default <Target> without folder creation or fallback' -ForEach @(@{ Target = 'store'; Code = 'default_store_unavailable' }, @{ Target = 'Inbox'; Code = 'default_inbox_unavailable' }) {
        $double = New-BrowseDouble
        if ($Target -eq 'store') { $double.state.application.methods.GetNamespace.properties.DefaultStore = $null }
        else { $double.state.application.methods.GetNamespace.properties.DefaultStore.methods.GetDefaultFolder = $null }
        $result = Browse $double
        $result.errors[0].code | Should -Be $Code
        Assert-NoBroadOrBodyReads $double
        Assert-AllReleased $double
    }
    It 'fails an unknown exact pair instead of falling back to Inbox' {
        $double = New-BrowseDouble
        $result = Browse $double (BrowseRequest ',"storeId":"s-a","folderId":"missing"')
        $result.status | Should -Be failed
        $result.errors[0].code | Should -Be folder_unavailable
        $double.state.calls | Should -Not -Contain 'session.DefaultStore'
        Assert-AllReleased $double
    }
    It 'fails mismatched resolved exact folder identity' {
        $double = New-BrowseDouble
        $double.state.folder.properties.EntryID = 'moved'
        $result = Browse $double (BrowseRequest ',"storeId":"s-a","folderId":"inbox"')
        $result.errors[0].code | Should -Be folder_identity_unavailable
        Assert-AllReleased $double
    }
    It 'retains exact source when display Store acquisition is denied' {
        $double = New-BrowseDouble -Failures @{ 'inbox.Store' = $true }
        $result = Browse $double (BrowseRequest ',"storeId":"s-a","folderId":"inbox"')
        $result.status | Should -Be partial
        $result.data.browse.returnedCount | Should -Be 10
        $result.data.browse.summaries[0].storeId | Should -Be s-a
        $result.data.browse.summaries[0].storeName | Should -BeNullOrEmpty
        Assert-AllReleased $double
    }
}

Describe 'Browse cooperative budgets, cancellation and safe progress' {
    It 'stops after a blocking call exhausts the budget without starting another COM call' {
        $double = New-BrowseDouble
        $time = @{ now = 0 }
        $double.state.actions['browse-999.Subject'] = { $time.now = 30 }.GetNewClosure()
        $clock = { $time.now }.GetNewClosure()
        $result = Browse $double (BrowseRequest) { $false } $clock
        $result.status | Should -Be partial
        $result.data.browse.stopReason | Should -Be time_budget
        $result.data.browse.returnedCount | Should -Be 1
        $double.state.calls | Should -Not -Contain 'browse-999.Parent'
        $result.data.browse.scanComplete | Should -BeFalse
        Assert-AllReleased $double
    }
    It 'checks a time budget after connection and before namespace acquisition' {
        $double = New-BrowseDouble
        $time = @{ now = 0 }; $connect = $double.adapter.Connect
        $double.adapter.Connect = { $value = & $connect; $time.now = 30; return $value }.GetNewClosure()
        $result = Browse $double (BrowseRequest) { $false } { $time.now }.GetNewClosure()
        $result.status | Should -Be partial
        $result.data.browse.stopReason | Should -Be time_budget
        $double.state.calls | Should -Not -Contain 'application.GetNamespace'
        Assert-AllReleased $double
    }
    It 'cancels before connection with no mailbox access' {
        $double = New-BrowseDouble
        $result = Browse $double (BrowseRequest) { $true }
        $result.status | Should -Be cancelled
        (Get-MailLensExitCode $result.status) | Should -Be 130
        $double.state.connections | Should -Be 0
        Assert-AllReleased $double
    }
    It 'cancels between calls, retains attempted useful summaries and releases acquired references' {
        $double = New-BrowseDouble; $state = $double.state
        $cancel = { $state.calls.Contains('browse-999.Release') }.GetNewClosure()
        $result = Browse $double (BrowseRequest) $cancel
        $result.status | Should -Be cancelled
        $result.data.browse.stopReason | Should -Be cancelled
        $result.data.browse.returnedCount | Should -Be 2
        $result.data.browse.scanComplete | Should -BeFalse
        @($state.calls | Where-Object { $_ -match '^inbox-items.Item\.' }).Count | Should -Be 2
        Assert-AllReleased $double
    }
    It 'writes only fixed phase/count/elapsed text and throttles repeated phases' {
        $double = New-BrowseDouble
        $lines = New-Object 'System.Collections.Generic.List[string]'
        $time = @{ now = 0 }
        $double.state.actions['browse-998.Subject'] = { $time.now = 5 }.GetNewClosure()
        $progress = { param($line) [void]$lines.Add($line) }.GetNewClosure()
        $clock = { $time.now }.GetNewClosure()
        $null = Browse $double (BrowseRequest) { $false } $clock $progress
        $lines.Count | Should -Be 7
        foreach ($line in $lines) { $line | Should -Match '^MailLens phase=(connect|namespace|acquire-folder|acquire-items|sort-items|browse-items) count=\d+ elapsedSeconds=\d+$' }
        ($lines -join '|') | Should -Not -Match '合成|工作|s-a|inbox-items|browse-1000|Synthetic|Exception'
        Assert-AllReleased $double
    }
}

Describe 'New fields and lightweight diagnose' {
    It 'rejects <Extra> before COM' -ForEach @(
        @{ Extra = ',"storeId":"s-a"' }, @{ Extra = ',"folderId":"inbox"' },
        @{ Extra = ',"storeId":"","folderId":"inbox"' }, @{ Extra = ',"storeId":1,"folderId":"inbox"' },
        @{ Extra = ',"limit":0' }, @{ Extra = ',"limit":101' }, @{ Extra = ',"limit":1.5' }, @{ Extra = ',"limit":"10"' }, @{ Extra = ',"limit":true' },
        @{ Extra = ',"maxItems":0' }, @{ Extra = ',"maxItems":10001' }, @{ Extra = ',"maxItems":9' }, @{ Extra = ',"maxItems":null' },
        @{ Extra = ',"maxDurationSeconds":0' }, @{ Extra = ',"maxDurationSeconds":301' }, @{ Extra = ',"maxDurationSeconds":"30"' },
        @{ Extra = ',"keywords":["synthetic"]' }, @{ Extra = ',"allHistory":true' }, @{ Extra = ',"mailboxes":["s-a"]' }, @{ Extra = ',"includeDescendants":true' }
    ) {
        $double = New-BrowseDouble
        $result = Browse $double (BrowseRequest $Extra)
        $result.status | Should -Be failed
        $double.state.connections | Should -Be 0
        Assert-AllReleased $double
    }
    It 'accepts explicit maximum and minimum budgets' -ForEach @(@{ Extra = ',"limit":1,"maxItems":1,"maxDurationSeconds":1' }, @{ Extra = ',"limit":100,"maxItems":10000,"maxDurationSeconds":300' }) {
        $double = New-BrowseDouble -Count 0
        (Browse $double (BrowseRequest $Extra)).status | Should -Be complete
        Assert-AllReleased $double
    }
    It 'rejects an unsupported browse runtime before connecting' {
        $double = New-BrowseDouble; $runtime = $native.Clone(); $runtime.isWindows = $false
        $result = Invoke-MailLensRequest (BrowseRequest) $double.adapter $runtime -Progress {}
        $result.errors[0].code | Should -Be unsupported_platform
        $double.state.connections | Should -Be 0
    }
    It 'diagnoses a connection without default folders, catalog, Items or mail reads' {
        $double = New-BrowseDouble -Failures @{ 'session.Stores' = $true; 'session.DefaultStore' = $true; 'inbox.Items' = $true }
        $request = [pscustomobject]@{ schemaVersion = [int]2; operation = 'diagnose' }
        $result = Invoke-MailLensRequest $request $double.adapter $native -Progress {}
        $result.status | Should -Be complete
        $result.data.outlookVersion | Should -BeExactly '16.0'
        $result.data.discoveryAttempted | Should -BeFalse
        $result.data.coverage.evaluated | Should -BeFalse
        $result.data.coverage.complete | Should -BeNullOrEmpty
        ($double.state.calls -join '|') | Should -BeExactly 'Connect|application.GetNamespace|application.Version|session.Release|application.Release'
        Assert-AllReleased $double
    }
    It 'retains diagnosis after optional version denial and never logs raw errors' {
        $double = New-BrowseDouble -Failures @{ 'application.Version' = $true }
        $result = Invoke-MailLensRequest ([pscustomobject]@{schemaVersion=[int]2;operation='diagnose'}) $double.adapter $native -Progress {}
        $result.status | Should -Be partial
        $result.data.discoveryAttempted | Should -BeFalse
        $result.warnings[0].code | Should -Be version_unavailable
        $result.warnings[0].message | Should -Not -Match 'Synthetic|private'
        Assert-AllReleased $double
    }
}

Describe 'Final browse call boundaries and source relocation' {
    It 'marks cancellation after the final date read even when the browse limit is fulfilled' {
        $double = New-BrowseDouble; $state = $double.state
        $result = Browse $double (BrowseRequest ',"limit":1') { $state.calls.Contains('browse-1000.ReceivedTime') }.GetNewClosure()
        $result.status | Should -Be cancelled
        $result.data.browse.stopReason | Should -Be cancelled
        $result.data.browse.returnedCount | Should -Be 1
        $result.data.browse.scanComplete | Should -BeFalse
        Assert-AllReleased $double
    }
    It 'discloses an over-budget final date read without dropping usable metadata' {
        $double = New-BrowseDouble; $time = @{ now = 0 }
        $double.state.actions['browse-1000.ReceivedTime'] = { $time.now = 30 }.GetNewClosure()
        $result = Browse $double (BrowseRequest ',"limit":1') { $false } { $time.now }.GetNewClosure()
        $result.status | Should -Be partial
        $result.data.browse.stopReason | Should -Be time_budget
        $result.data.browse.returnedCount | Should -Be 1
        Assert-AllReleased $double
    }
    It 'runs the copied module after source relocation' {
        $copy = Join-Path $TestDrive 'relocated-skill'
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../../skills/mail-lens') -Destination $copy -Recurse
        Import-Module (Join-Path $copy 'scripts/MailLens.Outlook.psm1') -Force
        try {
            $double = New-BrowseDouble
            (Browse $double).data.browse.returnedCount | Should -Be 10
            Assert-AllReleased $double
        } finally { Import-Module (Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/MailLens.Outlook.psm1') -Force }
    }
}

Describe 'Native browse atomic publication' -Tag Native {
    It 'publishes Unicode browse JSON from an input <Bom> BOM' -ForEach @(@{ Bom = $false }, @{ Bom = $true }) {
        $dir = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N')); [void][IO.Directory]::CreateDirectory($dir)
        $request = Join-Path $dir 'browse.request.json'; $output = Join-Path $dir 'browse.result.json'
        [IO.File]::WriteAllText($request, '{"schemaVersion":2,"operation":"browse","limit":1}', (New-Object Text.UTF8Encoding($Bom, $true)))
        $double = New-BrowseDouble
        (Invoke-MailLensFileRequest $request $output $double.adapter $native) | Should -Be 0
        $bytes = [IO.File]::ReadAllBytes($output); $bytes[0] | Should -Be 123
        $response = ConvertFrom-Json ((New-Object Text.UTF8Encoding($false, $true)).GetString($bytes))
        $response.data.browse.summaries[0].subject | Should -BeExactly '合成📮 1000'
        $response.data.browse.summaries[0].entryId | Should -BeExactly 'browse-1000'
        $response.data.discoveryAttempted | Should -BeFalse
        @(Get-ChildItem $dir -Force).Count | Should -Be 2
        { Invoke-MailLensFileRequest $request $output $double.adapter $native } | Should -Throw
        Assert-AllReleased $double
    }
    It 'publishes cancellation with attempted summaries and exit 130' {
        $dir = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N')); [void][IO.Directory]::CreateDirectory($dir)
        $request = Join-Path $dir 'browse.request.json'; $output = Join-Path $dir 'browse.result.json'
        [IO.File]::WriteAllText($request, '{"schemaVersion":2,"operation":"browse"}', (New-Object Text.UTF8Encoding($false)))
        $double = New-BrowseDouble; $state = $double.state
        $cancel = { $state.calls.Contains('browse-999.Release') }.GetNewClosure()
        (Invoke-MailLensFileRequest $request $output $double.adapter $native $cancel) | Should -Be 130
        $response = ConvertFrom-Json ([IO.File]::ReadAllText($output))
        $response.status | Should -Be cancelled
        $response.data.browse.returnedCount | Should -Be 2
        @(Get-ChildItem $dir -Force).Count | Should -Be 2
        Assert-AllReleased $double
    }
}
