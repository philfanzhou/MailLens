# MailLens schema 1. All Outlook access is read-only and routed through the adapter.
Set-StrictMode -Version 2.0

function Stop-MailLens {
    param([string]$Code, [string]$Message)
    $exception = New-Object System.InvalidOperationException($Message)
    $exception.Data['MailLensCode'] = $Code
    throw $exception
}

function New-MailLensComAdapter {
    # This seam supports synthetic tests; it is not a capability sandbox.
    @{
        Connect = {
            $application = $null
            try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
            catch {
                # Only a missing running instance permits starting Classic Outlook.
                if ($_.Exception.GetBaseException().HResult -ne -2147221021) { throw }
            }
            if ($null -ne $application) { return @{ Application = $application; Mode = 'existing' } }
            $application = New-Object -ComObject Outlook.Application -ErrorAction Stop
            return @{ Application = $application; Mode = 'started' }
        }
        Property = { param($Reference, [string]$Name) return ,($Reference.$Name) }
        Method = {
            param($Reference, [string]$Name, [object[]]$Arguments)
            switch ($Name) {
                'GetNamespace' { return ,($Reference.GetNamespace('MAPI')) }
                'GetRootFolder' { return ,($Reference.GetRootFolder()) }
                'GetFolderFromID' { return ,($Reference.GetFolderFromID([string]$Arguments[0], [string]$Arguments[1])) }
                'GetDefaultFolder' { return ,($Reference.GetDefaultFolder([int]$Arguments[0])) }
                'Item' { return ,($Reference.Item([int]$Arguments[0])) }
                default { throw 'Unsupported adapter method.' }
            }
        }
        Release = {
            param($Reference)
            if ($null -ne $Reference -and [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
                # Release only this acquisition, never FinalReleaseComObject on a shared RCW.
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Reference)
            }
        }
    }
}

function Get-MailLensRuntime {
    @{
        isWindows = ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT)
        powerShellVersion = $PSVersionTable.PSVersion.ToString()
        edition = $(if ($PSVersionTable.ContainsKey('PSEdition')) { $PSVersionTable.PSEdition } else { 'Desktop' })
        apartmentState = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
        timeZone = [TimeZoneInfo]::Local.Id
    }
}

function Assert-MailLensRequest {
    param($Request)
    if ($null -eq $Request -or $Request -isnot [pscustomobject]) {
        Stop-MailLens 'invalid_request' 'The request must be a JSON object.'
    }
    $names = @($Request.PSObject.Properties.Name)
    if ('schemaVersion' -notin $names -or $Request.schemaVersion -isnot [int] -or $Request.schemaVersion -ne 1) {
        Stop-MailLens 'invalid_request' 'schemaVersion must be the integer 1.'
    }
    if ('operation' -notin $names -or $Request.operation -isnot [string] -or
        $Request.operation -cnotin @('diagnose', 'folders', 'search', 'get')) {
        Stop-MailLens 'unsupported_operation' 'operation must be diagnose, folders, search or get.'
    }
    # Detail remains outside this search slice.
    if ($Request.operation -eq 'get') {
        Stop-MailLens 'unsupported_operation' 'This operation is not implemented in this version.'
    }
    $allowed = @('schemaVersion', 'operation')
    if ($Request.operation -in @('folders', 'search')) { $allowed += @('mailboxes', 'folders', 'includeDescendants') }
    if ($Request.operation -eq 'search') { $allowed += @('fromInclusive', 'toExclusive', 'allHistory', 'keywords', 'keywordGroups', 'listOnly', 'sender', 'summaryLimit') }
    foreach ($name in $names) {
        if ($name -cnotin $allowed) { Stop-MailLens 'invalid_request' 'The request contains an inapplicable or unknown field.' }
    }
    foreach ($field in @('mailboxes', 'folders')) {
        if ($field -in $names) {
            $values = $Request.$field
            if ($values -isnot [array] -or $values.Count -eq 0) {
                Stop-MailLens 'invalid_request' 'Selection fields must be nonempty arrays.'
            }
            foreach ($value in $values) {
                if ($value -is [string] -and -not [string]::IsNullOrWhiteSpace($value)) { continue }
                if ($field -eq 'folders' -and $value -is [pscustomobject]) {
                    $keys = @($value.PSObject.Properties.Name)
                    if ($keys.Count -eq 2 -and 'storeId' -cin $keys -and 'folderId' -cin $keys -and
                        $value.storeId -is [string] -and -not [string]::IsNullOrWhiteSpace($value.storeId) -and
                        $value.folderId -is [string] -and -not [string]::IsNullOrWhiteSpace($value.folderId)) { continue }
                }
                Stop-MailLens 'invalid_request' 'Mailbox selectors must be exact strings; folder selectors must be exact paths or storeId/folderId pairs.'
            }
        }
    }
    if ($Request.operation -eq 'search') { Assert-MailLensSearchRequest $Request }
    if ('includeDescendants' -in $names -and $Request.includeDescendants -isnot [bool]) {
        Stop-MailLens 'invalid_request' 'includeDescendants must be a boolean.'
    }
}

function Add-MailLensProblem {
    param($Context, [string]$Code, [string]$Message, $Location = $null)
    [void]$Context.Problems.Add([ordered]@{ code = $Code; message = $Message; location = $Location })
}

function Close-MailLensReference {
    param($Adapter, $Reference, $Context)
    if ($null -ne $Reference) {
        try { & $Adapter.Release $Reference }
        catch { Add-MailLensProblem $Context 'release_failed' 'An acquired Outlook reference could not be released.' }
    }
}

function Assert-MailLensNotCancelled {
    param($CancellationRequested)
    if (& $CancellationRequested) { Stop-MailLens 'cancelled' 'The operation was cancelled; discovery coverage is incomplete.' }
}

function Read-MailLensFolder {
    param($Adapter, $Folder, $StoreRow, $ParentId, $Context, $CancellationRequested)
    Assert-MailLensNotCancelled $CancellationRequested
    $row = [ordered]@{
        storeId = $StoreRow.storeId; storeName = $StoreRow.storeName
        folderId = $null; folderPath = $null; parentFolderId = $ParentId
        defaultItemType = $null; readable = $false; childrenReadable = $false
    }
    # Preserve each obtainable identity even when a later property is denied.
    foreach ($property in @('EntryID', 'FolderPath', 'DefaultItemType')) {
        try {
            $value = & $Adapter.Property $Folder $property
            switch ($property) {
                'EntryID' { $row.folderId = [string]$value }
                'FolderPath' { $row.folderPath = [string]$value }
                'DefaultItemType' { $row.defaultItemType = [int]$value }
            }
        } catch { Add-MailLensProblem $Context 'folder_denied' 'A folder property is unavailable.' $row }
    }
    if (-not [string]::IsNullOrEmpty($row.folderId)) {
        $key = $StoreRow.storeId.Length.ToString() + ':' + $StoreRow.storeId + $row.folderId
        if (-not $Context.Seen.Add($key)) { return }
    } else {
        # Without an identity, descending could loop forever on a cyclic provider.
        [void]$Context.Folders.Add($row)
        return
    }
    $row.readable = ($null -ne $row.folderPath -and $null -ne $row.defaultItemType)
    [void]$Context.Folders.Add($row)
    $children = $null
    try {
        Assert-MailLensNotCancelled $CancellationRequested
        $children = & $Adapter.Property $Folder 'Folders'
        $count = [int](& $Adapter.Property $children 'Count')
        $row.childrenReadable = $true
        for ($index = 1; $index -le $count; $index++) {
            Assert-MailLensNotCancelled $CancellationRequested
            $child = $null
            try {
                $child = & $Adapter.Method $children 'Item' @($index)
                Read-MailLensFolder $Adapter $child $StoreRow $row.folderId $Context $CancellationRequested
            } catch {
                if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
                $row.childrenReadable = $false
                Add-MailLensProblem $Context 'folder_denied' 'A child folder is unavailable.' ([ordered]@{ storeId = $StoreRow.storeId; parentFolderId = $row.folderId; childIndex = $index })
            } finally { Close-MailLensReference $Adapter $child $Context }
        }
    } catch {
        if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
        Add-MailLensProblem $Context 'folder_denied' 'The child-folder collection is unavailable.' $row
    } finally { Close-MailLensReference $Adapter $children $Context }
}

function Read-MailLensCatalog {
    param($Adapter, $Session, $Context, $CancellationRequested)
    $stores = $null
    try {
        $stores = & $Adapter.Property $Session 'Stores'
        $count = [int](& $Adapter.Property $stores 'Count')
        $Context.StoreCount = $count
        for ($index = 1; $index -le $count; $index++) {
            Assert-MailLensNotCancelled $CancellationRequested
            $store = $null; $root = $null
            $row = [ordered]@{ storeId = $null; storeName = $null; readable = $false; index = $index }
            [void]$Context.Stores.Add($row)
            try {
                $store = & $Adapter.Method $stores 'Item' @($index)
                $row.storeId = [string](& $Adapter.Property $store 'StoreID')
                $row.storeName = [string](& $Adapter.Property $store 'DisplayName')
                if ([string]::IsNullOrEmpty($row.storeId)) { throw 'Missing store identity.' }
                $root = & $Adapter.Method $store 'GetRootFolder' @()
                $row.readable = $true
                Read-MailLensFolder $Adapter $root $row $null $Context $CancellationRequested
            } catch {
                if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
                Add-MailLensProblem $Context 'store_denied' 'A mounted store or its root folder is unavailable.' $row
            } finally {
                Close-MailLensReference $Adapter $root $Context
                Close-MailLensReference $Adapter $store $Context
            }
        }
    } finally { Close-MailLensReference $Adapter $stores $Context }
}

function Resolve-MailLensScope {
    param($Request, $Context)
    $stores = @($Context.Stores.ToArray())
    $names = @($Request.PSObject.Properties.Name)
    if ('mailboxes' -in $names) {
        $chosen = New-Object 'System.Collections.Generic.List[object]'
        foreach ($selector in $Request.mailboxes) {
            $matches = @($stores | Where-Object { [string]::Equals($_.storeId, $selector, [StringComparison]::Ordinal) })
            if ($matches.Count -eq 0) {
                $matches = @($stores | Where-Object { [string]::Equals($_.storeName, $selector, [StringComparison]::Ordinal) })
            }
            if ($matches.Count -eq 0) { Stop-MailLens 'unknown_selection' 'A mailbox selector does not match a mounted store. Inspect the catalog and use its exact storeId.' }
            if ($matches.Count -ne 1) { Stop-MailLens 'ambiguous_selection' 'A mailbox display name is ambiguous. Inspect the catalog choices and use an exact storeId.' }
            if ($chosen -notcontains $matches[0]) { [void]$chosen.Add($matches[0]) }
        }
        $stores = @($chosen.ToArray())
    }
    $storeIds = @($stores | ForEach-Object { $_.storeId })
    $folders = @($Context.Folders.ToArray() | Where-Object { $_.storeId -cin $storeIds })
    if ('folders' -in $names) {
        $chosen = New-Object 'System.Collections.Generic.List[object]'
        foreach ($selector in $Request.folders) {
            $matches = @($folders | Where-Object {
                if ($selector -is [string]) { [string]::Equals($_.folderPath, $selector, [StringComparison]::Ordinal) }
                else { [string]::Equals($_.storeId, $selector.storeId, [StringComparison]::Ordinal) -and [string]::Equals($_.folderId, $selector.folderId, [StringComparison]::Ordinal) }
            })
            if ($matches.Count -eq 0) { Stop-MailLens 'unknown_selection' 'A folder selector does not match the resolved stores. Inspect the catalog and use a storeId/folderId pair.' }
            if ($matches.Count -ne 1) { Stop-MailLens 'ambiguous_selection' 'A folder path is ambiguous. Inspect the catalog choices and use a storeId/folderId pair.' }
            if ($chosen -notcontains $matches[0]) { [void]$chosen.Add($matches[0]) }
        }
        $descendants = ('includeDescendants' -notin $names -or $Request.includeDescendants)
        if ($descendants) {
            # Follow parent identities rather than interpreting display-path prefixes.
            for ($index = 0; $index -lt $chosen.Count; $index++) {
                $parent = $chosen[$index]
                foreach ($candidate in $folders) {
                    if ($candidate.storeId -ceq $parent.storeId -and $candidate.parentFolderId -ceq $parent.folderId -and $chosen -notcontains $candidate) { [void]$chosen.Add($candidate) }
                }
            }
        }
        $folders = @($chosen.ToArray())
        # An explicit folder selection narrows the store list as well.
        $selectedIds = @($folders | ForEach-Object { $_.storeId })
        $stores = @($stores | Where-Object { $_.storeId -cin $selectedIds })
    }
    return [ordered]@{ stores = @($stores); folders = @($folders) }
}


function Convert-MailLensBound {
    param($Value)
    if ($Value -isnot [string] -or $Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,7})?(Z|[+-]\d{2}:\d{2})$') {
        Stop-MailLens 'invalid_request' 'Date bounds must be full ISO 8601 timestamps with explicit offsets.'
    }
    $parsed = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        Stop-MailLens 'invalid_request' 'A date bound is invalid.'
    }
    return $parsed.ToUniversalTime()
}

function Assert-MailLensSearchRequest {
    param($Request)
    $names = @($Request.PSObject.Properties.Name)
    foreach ($field in @('allHistory', 'listOnly')) {
        if ($field -in $names -and $Request.$field -isnot [bool]) { Stop-MailLens 'invalid_request' 'Search mode flags must be booleans.' }
    }
    $bounds = @($names | Where-Object { $_ -cin @('fromInclusive', 'toExclusive') })
    if ($bounds.Count -ne 0 -and $bounds.Count -ne 2) { Stop-MailLens 'invalid_request' 'Supply both date bounds.' }
    if ('allHistory' -in $names -and $bounds.Count -gt 0) { Stop-MailLens 'invalid_request' 'allHistory and explicit date bounds are mutually exclusive.' }
    if ($bounds.Count -eq 2) {
        $from = Convert-MailLensBound $Request.fromInclusive
        $to = Convert-MailLensBound $Request.toExclusive
        if ($from -ge $to) { Stop-MailLens 'invalid_request' 'fromInclusive must precede toExclusive.' }
    }
    $modes = @($names | Where-Object { $_ -cin @('keywords', 'keywordGroups', 'listOnly') })
    if ($modes.Count -ne 1 -or ('listOnly' -in $names -and -not $Request.listOnly)) {
        Stop-MailLens 'invalid_request' 'Choose nonempty keywords, keywordGroups or listOnly true.'
    }
    foreach ($field in @('keywords', 'keywordGroups')) {
        if ($field -notin $names) { continue }
        $values = $Request.$field
        if ($values -isnot [array] -or $values.Count -eq 0) { Stop-MailLens 'invalid_request' 'Keyword modes require nonempty arrays.' }
        foreach ($value in $values) {
            if ($field -eq 'keywords' -and $value -isnot [string]) { Stop-MailLens 'invalid_request' 'Every keyword must be a literal string.' }
            $terms = @($value)
            if ($field -eq 'keywordGroups') {
                if ($value -isnot [array] -or $value.Count -eq 0) { Stop-MailLens 'invalid_request' 'Each keyword group must be a nonempty array.' }
                $terms = $value
            }
            foreach ($term in $terms) {
                if ($term -isnot [string] -or [string]::IsNullOrWhiteSpace($term)) { Stop-MailLens 'invalid_request' 'Every keyword must be a nonempty literal string.' }
            }
        }
    }
    if ('sender' -in $names -and ($Request.sender -isnot [string] -or [string]::IsNullOrWhiteSpace($Request.sender))) {
        Stop-MailLens 'invalid_request' 'sender must be a nonempty literal string.'
    }
    if ('summaryLimit' -in $names -and ($Request.summaryLimit -isnot [int] -or $Request.summaryLimit -lt 1)) {
        Stop-MailLens 'invalid_request' 'summaryLimit must be a positive integer.'
    }
}

function New-MailLensSearchData {
    param($Request, [DateTimeOffset]$StartedAt, $Runtime)
    $names = @($Request.PSObject.Properties.Name)
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Runtime.timeZone)
    $from = $StartedAt.ToUniversalTime().AddDays(-30); $to = $StartedAt.ToUniversalTime()
    $mode = 'rolling30Days'
    if ('fromInclusive' -in $names) {
        $from = Convert-MailLensBound $Request.fromInclusive; $to = Convert-MailLensBound $Request.toExclusive; $mode = 'explicit'
    }
    if ('allHistory' -in $names -and $Request.allHistory) { $from = $null; $to = $null; $mode = 'allHistory' }
    $data = [ordered]@{
        range = [ordered]@{ mode = $mode; timeZone = $zone.Id; fromInclusive = $null; toExclusive = $null; localFromInclusive = $null; localToExclusive = $null }
        keywords = @(); keywordGroups = @(); listOnly = ('listOnly' -in $names); sender = $null
        summaryLimit = 100; matchedCount = 0; returnedCount = 0; resultsTruncated = $false
        summaries = @(); folderCoverage = New-Object 'System.Collections.Generic.List[object]'
    }
    if ($null -ne $from) {
        $data.range.fromInclusive = $from.ToString('o'); $data.range.toExclusive = $to.ToString('o')
        $data.range.localFromInclusive = [TimeZoneInfo]::ConvertTime($from, $zone).ToString('o')
        $data.range.localToExclusive = [TimeZoneInfo]::ConvertTime($to, $zone).ToString('o')
    }
    if ('keywords' -in $names) { $data.keywords = @($Request.keywords) }
    if ('keywordGroups' -in $names) { $data.keywordGroups = @($Request.keywordGroups) }
    if ('sender' -in $names) { $data.sender = $Request.sender }
    if ('summaryLimit' -in $names) { $data.summaryLimit = $Request.summaryLimit }
    return @{ data = $data; zone = $zone; from = $from; to = $to; matches = New-Object 'System.Collections.Generic.List[object]' }
}

function Read-MailLensSentFolders {
    param($Adapter, $Session, $Scope, $Context, $CancellationRequested)
    $sent = @{}; $stores = $null
    try {
        $stores = & $Adapter.Property $Session 'Stores'
        foreach ($row in $Scope.stores) {
            Assert-MailLensNotCancelled $CancellationRequested
            $store = $null; $folder = $null
            try {
                $store = & $Adapter.Method $stores 'Item' @($row.index)
                if ([string](& $Adapter.Property $store 'StoreID') -cne $row.storeId) { throw 'Store changed.' }
                $folder = & $Adapter.Method $store 'GetDefaultFolder' @(5)
                $ids = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
                if ($null -ne $folder) {
                    $id = [string](& $Adapter.Property $folder 'EntryID')
                    if ([string]::IsNullOrEmpty($id)) { throw 'Missing Sent folder identity.' }
                    [void]$ids.Add($id)
                    # Follow the full catalog even when only a Sent descendant was selected.
                    do {
                        $added = $false
                        foreach ($candidate in $Context.Folders) {
                            if ($candidate.storeId -ceq $row.storeId -and $null -ne $candidate.parentFolderId -and $ids.Contains($candidate.parentFolderId)) {
                                if ($ids.Add($candidate.folderId)) { $added = $true }
                            }
                        }
                    } while ($added)
                }
                $sent[$row.storeId] = $ids
            } catch {
                if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
                Add-MailLensProblem $Context 'sent_folder_denied' 'The store Sent Items identity is unavailable; sent-mail date classification is incomplete.' $row
            } finally {
                Close-MailLensReference $Adapter $folder $Context
                Close-MailLensReference $Adapter $store $Context
            }
        }
    } finally { Close-MailLensReference $Adapter $stores $Context }
    return $sent
}

function Convert-MailLensMailTime {
    param($Value, $Zone)
    # Outlook's unset date sentinel is year 4501. Do not parse locale text as dates.
    if ($Value -isnot [DateTime] -or $Value -le [DateTime]'1601-01-01' -or $Value.Year -ge 4501) { return $null }
    if ($Value.Kind -eq [DateTimeKind]::Utc) { return [DateTimeOffset]::new($Value) }
    $local = [DateTime]::SpecifyKind($Value, [DateTimeKind]::Unspecified)
    if ($Zone.IsInvalidTime($local) -or $Zone.IsAmbiguousTime($local)) {
        Stop-MailLens 'date_dst_gap' 'An Outlook local timestamp is ambiguous or nonexistent in the Windows timezone.'
    }
    return [DateTimeOffset]::new([TimeZoneInfo]::ConvertTimeToUtc($local, $Zone))
}

function Read-MailLensEffectiveTime {
    param($Adapter, $Item, $IsSentFolder, $SentKnown, $Zone)
    $isSent = & $Adapter.Property $Item 'Sent'
    if ($isSent -isnot [bool]) { throw 'Invalid Sent flag.' }
    if (-not $isSent) { $fields = @('CreationTime') }
    elseif (-not $SentKnown) { throw 'Sent classification unavailable.' }
    elseif ($IsSentFolder) { $fields = @('SentOn') }
    else { $fields = @('ReceivedTime', 'SentOn', 'CreationTime') }
    foreach ($field in $fields) {
        $value = Convert-MailLensMailTime (& $Adapter.Property $Item $field) $Zone
        if ($null -ne $value) { return @{ value = $value; field = $field } }
    }
    Stop-MailLens 'date_unavailable' 'No valid effective timestamp is available for this mail item.'
}

function Test-MailLensTerms {
    param([string]$Subject, [string]$Body, $Data)
    if ($Data.listOnly) { return $true }
    $groups = $Data.keywordGroups
    if ($Data.keywords.Count -gt 0) { $groups = @(,$Data.keywords) }
    foreach ($group in $groups) {
        $matched = $false
        foreach ($term in $group) {
            if ($Subject.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or $Body.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $matched = $true; break }
        }
        if (-not $matched) { return $false }
    }
    return $true
}

function Assert-MailLensItemLocation {
    param($Adapter, $Item, $Row, $Context)
    $parent = $null
    try {
        $parent = & $Adapter.Property $Item 'Parent'
        if ([string](& $Adapter.Property $parent 'EntryID') -cne $Row.folderId -or
            [string](& $Adapter.Property $parent 'StoreID') -cne $Row.storeId) { throw 'The item moved during retrieval.' }
    } finally { Close-MailLensReference $Adapter $parent $Context }
}

function Read-MailLensSearch {
    param($Adapter, $Session, $Scope, $Context, $Search, $CancellationRequested)
    $sent = Read-MailLensSentFolders $Adapter $Session $Scope $Context $CancellationRequested
    foreach ($row in $Scope.folders) {
        Assert-MailLensNotCancelled $CancellationRequested
        $coverage = [ordered]@{ storeId = $row.storeId; folderId = $row.folderId; folderPath = $row.folderPath
            visitedCount = 0; mailItemCount = 0; inRangeCount = 0; matchedCount = 0; unreadableCount = 0; complete = $false }
        [void]$Search.data.folderCoverage.Add($coverage)
        $folder = $null; $items = $null
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        try {
            if (-not $row.readable) { throw 'Folder metadata unavailable.' }
            $folder = & $Adapter.Method $Session 'GetFolderFromID' @($row.folderId, $row.storeId)
            if ([string](& $Adapter.Property $folder 'EntryID') -cne $row.folderId -or
                [string](& $Adapter.Property $folder 'StoreID') -cne $row.storeId) { throw 'Folder identity changed.' }
            $items = & $Adapter.Property $folder 'Items'
            $count = [int](& $Adapter.Property $items 'Count')
            for ($index = 1; $index -le $count; $index++) {
                Assert-MailLensNotCancelled $CancellationRequested
                $item = $null
                $location = [ordered]@{ storeId = $row.storeId; folderId = $row.folderId; itemIndex = $index; entryId = $null }
                $coverage.visitedCount++
                try {
                    $item = & $Adapter.Method $items 'Item' @($index)
                    if ([int](& $Adapter.Property $item 'Class') -ne 43) { continue }
                    $coverage.mailItemCount++
                    $entryId = [string](& $Adapter.Property $item 'EntryID')
                    $location.entryId = $entryId
                    if ([string]::IsNullOrEmpty($entryId) -or -not $seen.Add($entryId)) { throw 'Missing or repeated identity during scan.' }
                    Assert-MailLensItemLocation $Adapter $item $row $Context
                    $time = Read-MailLensEffectiveTime $Adapter $item ($sent.ContainsKey($row.storeId) -and $sent[$row.storeId].Contains($row.folderId)) ($sent.ContainsKey($row.storeId)) $Search.zone
                    if ($null -ne $Search.from -and ($time.value -lt $Search.from -or $time.value -ge $Search.to)) { continue }
                    $coverage.inRangeCount++
                    $modified = & $Adapter.Property $item 'LastModificationTime'
                    $subject = & $Adapter.Property $item 'Subject'
                    $body = & $Adapter.Property $item 'Body'
                    if ($subject -isnot [string] -or $body -isnot [string]) { throw 'Plain subject/body unavailable.' }
                    $sender = $null; $senderMatches = $true
                    if ($null -ne $Search.data.sender) {
                        $sender = [ordered]@{ displayName = $null; rawAddress = $null; addressType = $null }
                        $available = 0; $senderGap = $false
                        foreach ($pair in @(@('SenderName', 'displayName'), @('SenderEmailAddress', 'rawAddress'), @('SenderEmailType', 'addressType'))) {
                            try {
                                $value = & $Adapter.Property $item $pair[0]
                                if ($value -is [string]) { $sender[$pair[1]] = $value; if ($pair[1] -ne 'addressType') { $available++ } }
                                else { throw 'Sender property unavailable.' }
                            } catch { $senderGap = $true; Add-MailLensProblem $Context 'sender_denied' 'Sender metadata is partly unavailable; narrowing coverage is incomplete.' $location }
                        }
                        if ($available -eq 0) { throw 'Sender metadata unavailable.' }
                        if ($senderGap) { $coverage.unreadableCount++ }
                        $senderMatches = ([string]$sender.displayName).IndexOf($Search.data.sender, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                            ([string]$sender.rawAddress).IndexOf($Search.data.sender, [StringComparison]::OrdinalIgnoreCase) -ge 0
                    }
                    Assert-MailLensNotCancelled $CancellationRequested
                    Assert-MailLensItemLocation $Adapter $item $row $Context
                    if ([string](& $Adapter.Property $item 'EntryID') -cne $entryId -or
                        (& $Adapter.Property $item 'LastModificationTime') -ne $modified) { throw 'The item changed during retrieval.' }
                    if ($senderMatches -and (Test-MailLensTerms $subject $body $Search.data)) {
                        $coverage.matchedCount++; $Search.data.matchedCount++
                        $previewLength = [Math]::Min(500, $body.Length)
                        if ($previewLength -gt 0 -and [char]::IsHighSurrogate($body[$previewLength - 1])) { $previewLength-- }
                        [void]$Search.matches.Add([ordered]@{ storeId = $row.storeId; storeName = $row.storeName
                            folderId = $row.folderId; folderPath = $row.folderPath; entryId = $entryId; subject = $subject
                            timestamp = $time.value.ToUniversalTime().ToString('o'); timestampField = $time.field
                            preview = $body.Substring(0, $previewLength); previewSource = 'Body'; sender = $sender })
                    }
                } catch {
                    if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
                    $coverage.unreadableCount++
                    $code = $_.Exception.Data['MailLensCode']; if ($null -eq $code) { $code = 'item_unavailable' }
                    Add-MailLensProblem $Context $code 'A mail item could not be evaluated reliably; search coverage is incomplete.' $location
                } finally { Close-MailLensReference $Adapter $item $Context }
            }
            if ([int](& $Adapter.Property $items 'Count') -ne $count) { throw 'The item collection changed during scan.' }
            $coverage.complete = ($coverage.unreadableCount -eq 0)
        } catch {
            if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
            $coverage.unreadableCount++
            Add-MailLensProblem $Context 'search_folder_denied' 'The selected folder or its item collection is unavailable or changed; search coverage is incomplete.' $row
        } finally {
            Close-MailLensReference $Adapter $items $Context
            Close-MailLensReference $Adapter $folder $Context
        }
    }
}

function Complete-MailLensSearchData {
    param($Search)
    $comparison = [System.Comparison[object]]{
        param($left, $right)
        $order = [StringComparer]::Ordinal.Compare($right.timestamp, $left.timestamp)
        if ($order -ne 0) { return $order }
        foreach ($field in @('storeId', 'folderId', 'entryId')) {
            $order = [StringComparer]::Ordinal.Compare($left[$field], $right[$field])
            if ($order -ne 0) { return $order }
        }
        return 0
    }
    $Search.matches.Sort($comparison)
    $Search.data.returnedCount = [Math]::Min($Search.data.summaryLimit, $Search.matches.Count)
    $Search.data.summaries = @($Search.matches.ToArray() | Select-Object -First $Search.data.returnedCount)
    $Search.data.resultsTruncated = ($Search.data.matchedCount -gt $Search.data.returnedCount)
    $Search.data.folderCoverage = @($Search.data.folderCoverage.ToArray())
}

function Invoke-MailLensRequest {
    [CmdletBinding()]
    param($Request, $Adapter = (New-MailLensComAdapter), $Runtime = (Get-MailLensRuntime),
        [scriptblock]$CancellationRequested = { $false }, [DateTimeOffset]$StartedAt = [DateTimeOffset]::UtcNow)
    $context = @{
        Stores = New-Object 'System.Collections.Generic.List[object]'
        Folders = New-Object 'System.Collections.Generic.List[object]'
        Problems = New-Object 'System.Collections.Generic.List[object]'
        Seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        StoreCount = 0
    }
    $response = [ordered]@{
        schemaVersion = 1; operation = $null; status = 'failed'
        startedAt = $StartedAt.ToString('o'); endedAt = $null
        scope = [ordered]@{ stores = @(); folders = @() }; warnings = @(); errors = @()
        data = [ordered]@{ runtime = $Runtime; sessionMode = $null; stores = @(); folders = @(); coverage = $null }
    }
    $application = $null; $session = $null; $search = $null
    try {
        Assert-MailLensRequest $Request
        $response.operation = $Request.operation
        if (-not $Runtime.isWindows) { Stop-MailLens 'unsupported_platform' 'MailLens requires native Windows.' }
        if ($Runtime.edition -cne 'Desktop' -or $Runtime.powerShellVersion -notmatch '^5\.1\.') { Stop-MailLens 'unsupported_powershell' 'MailLens requires Windows PowerShell 5.1.' }
        if ($Runtime.apartmentState -cne 'STA') { Stop-MailLens 'unsupported_apartment' 'Launch powershell.exe with -NoProfile -STA.' }
        if ($Request.operation -eq 'search') {
            $search = New-MailLensSearchData $Request $StartedAt $Runtime
            $response.data.search = $search.data
        }
        Assert-MailLensNotCancelled $CancellationRequested
        try {
            $connection = & $Adapter.Connect
            $application = $connection.Application
            $response.data.sessionMode = [string]$connection.Mode
            $session = & $Adapter.Method $application 'GetNamespace' @('MAPI')
        } catch { Stop-MailLens 'outlook_unavailable' 'Classic Outlook COM or the current profile is unavailable; allow legitimate prompts or check Classic Outlook installation.' }
        Read-MailLensCatalog $Adapter $session $context $CancellationRequested
        $response.scope = Resolve-MailLensScope $Request $context
        if ($Request.operation -eq 'search') {
            Read-MailLensSearch $Adapter $session $response.scope $context $search $CancellationRequested
            Assert-MailLensNotCancelled $CancellationRequested
        }
        $response.status = 'complete'
        if ($context.Problems.Count -gt 0) { $response.status = 'partial' }
        if ($context.StoreCount -gt 0 -and @($context.Stores.ToArray() | Where-Object { $_.readable }).Count -eq 0) { $response.status = 'failed' }
        if ($context.StoreCount -gt 0 -and @($context.Folders.ToArray() | Where-Object { $_.readable }).Count -eq 0) { $response.status = 'failed' }
        if ($response.scope.stores.Count -gt 0 -and @($response.scope.stores | Where-Object { $_.readable }).Count -eq 0) { $response.status = 'failed' }
    } catch {
        $code = $_.Exception.Data['MailLensCode']
        $message = $_.Exception.Message
        if ($null -eq $code) { $code = 'discovery_failed'; $message = 'Outlook discovery failed; coverage is incomplete.' }
        $response.errors = @([ordered]@{ code = [string]$code; message = $message })
        if ($code -eq 'cancelled') { $response.status = 'cancelled' } else { $response.status = 'failed' }
    } finally {
        Close-MailLensReference $Adapter $session $context
        Close-MailLensReference $Adapter $application $context
        if ($null -ne $search) { Complete-MailLensSearchData $search }
        $response.data.stores = @($context.Stores.ToArray())
        $response.data.folders = @($context.Folders.ToArray())
        $response.data.coverage = [ordered]@{
            mountedStores = $context.StoreCount; discoveredStores = $context.Stores.Count
            readableStores = @($context.Stores.ToArray() | Where-Object { $_.readable }).Count
            discoveredFolders = $context.Folders.Count
            readableFolders = @($context.Folders.ToArray() | Where-Object { $_.readable }).Count
            complete = ($response.status -eq 'complete' -and $context.Problems.Count -eq 0)
        }
        $response.warnings = @($context.Problems.ToArray())
        if ($response.status -eq 'complete' -and $context.Problems.Count -gt 0) { $response.status = 'partial' }
        $response.endedAt = [DateTimeOffset]::UtcNow.ToString('o')
    }
    return $response
}

function Get-MailLensExitCode {
    param([string]$Status)
    switch ($Status) { 'complete' { 0 }; 'partial' { 2 }; 'cancelled' { 130 }; default { 1 } }
}

function Invoke-MailLensFileRequest {
    [CmdletBinding()]
    param([string]$RequestPath, [string]$OutputPath, $Adapter = (New-MailLensComAdapter),
        $Runtime = (Get-MailLensRuntime), [scriptblock]$CancellationRequested = { $false })
    $started = [DateTimeOffset]::UtcNow
    $reservation = $null; $temporary = $null; $lockPath = $null
    try {
        foreach ($path in @($RequestPath, $OutputPath)) {
            if (-not [IO.Path]::IsPathRooted($path) -or $path -match '^[A-Za-z]:($|[^\\/])' -or $path -match '^\\[^\\]') { throw 'RequestPath and OutputPath must be absolute file paths.' }
        }
        $requestFile = [IO.Path]::GetFullPath($RequestPath)
        $outputFile = [IO.Path]::GetFullPath($OutputPath)
        if ([string]::Equals($requestFile, $outputFile, [StringComparison]::OrdinalIgnoreCase)) { throw 'Request and output paths must be different.' }
        if ([IO.File]::Exists($outputFile) -or [IO.Directory]::Exists($outputFile)) { throw 'OutputPath already exists; use a unique result file.' }
        $parent = [IO.Path]::GetDirectoryName($outputFile)
        if (-not [IO.Directory]::Exists($parent)) { throw 'The output directory must already exist.' }
        $lockPath = $outputFile + '.maillens-lock'
        $reservation = [IO.File]::Open($lockPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        # Validate writable output before COM. CreateNew also refuses a temp collision.
        $temporary = [IO.Path]::Combine($parent, '.maillens-' + [Guid]::NewGuid().ToString('N') + '.tmp')
        $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $stream.Dispose()
    } catch {
        if ($null -ne $reservation) { $reservation.Dispose(); [IO.File]::Delete($lockPath) }
        throw 'Request/output paths are unavailable, invalid or already reserved. Use absolute paths and a unique writable output.'
    }
    try {
        $response = $null
        try {
            $encoding = New-Object Text.UTF8Encoding($false, $true)
            # Decode bytes ourselves: StreamReader would silently accept UTF-16 BOMs.
            $json = $encoding.GetString([IO.File]::ReadAllBytes($requestFile))
            if ($json.Length -gt 0 -and $json[0] -eq [char]0xFEFF) { $json = $json.Substring(1) }
            $request = ConvertFrom-Json -InputObject $json -ErrorAction Stop
            $response = Invoke-MailLensRequest $request $Adapter $Runtime $CancellationRequested $started
        } catch {
            $response = [ordered]@{
                schemaVersion = 1; operation = $null; status = 'failed'
                startedAt = $started.ToString('o'); endedAt = [DateTimeOffset]::UtcNow.ToString('o')
                scope = [ordered]@{ stores = @(); folders = @() }; warnings = @()
                errors = @([ordered]@{ code = 'invalid_request_file'; message = 'The UTF-8 request file is unreadable or contains malformed JSON.' })
                data = $null
            }
        }
        $serialized = ConvertTo-Json -InputObject $response -Depth 30 -Compress
        [IO.File]::WriteAllText($temporary, $serialized, (New-Object Text.UTF8Encoding($false)))
        # File.Move never replaces a destination, including a non-cooperating writer's file.
        [IO.File]::Move($temporary, $outputFile)
        return (Get-MailLensExitCode $response.status)
    } finally {
        if ($null -ne $temporary -and [IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        $reservation.Dispose()
        [IO.File]::Delete($lockPath)
    }
}

Export-ModuleMember -Function New-MailLensComAdapter, Get-MailLensRuntime, Invoke-MailLensRequest, Invoke-MailLensFileRequest, Get-MailLensExitCode
