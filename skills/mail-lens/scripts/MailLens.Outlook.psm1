# MailLens schema 2. All Outlook access is read-only and routed through the adapter.
Set-StrictMode -Version 2.0

function Stop-MailLens {
    param([string]$Code, [string]$Message, $Cause = $null)
    $exception = New-Object System.InvalidOperationException($Message)
    $exception.Data['MailLensCode'] = $Code
    if ($null -ne $Cause) { $exception.Data['MailLensHResult'] = $Cause.HResult }
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
                'GetItemFromID' { return ,($Reference.GetItemFromID([string]$Arguments[0], [string]$Arguments[1])) }
                'GetDefaultFolder' { return ,($Reference.GetDefaultFolder([int]$Arguments[0])) }
                'Item' { return ,($Reference.Item([int]$Arguments[0])) }
                'Restrict' { return ,($Reference.Restrict([string]$Arguments[0])) }
                'Sort' { $Reference.Sort([string]$Arguments[0], [bool]$Arguments[1]); return }
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
    if ('schemaVersion' -notin $names -or $Request.schemaVersion -isnot [int] -or $Request.schemaVersion -ne 2) {
        Stop-MailLens 'unsupported_schema' 'schemaVersion must be the integer 2; update the complete package and request together.'
    }
    if ('operation' -notin $names -or $Request.operation -isnot [string] -or
        $Request.operation -cnotin @('diagnose', 'browse', 'folders', 'search', 'get')) {
        Stop-MailLens 'unsupported_operation' 'operation must be diagnose, browse, folders, search or get.'
    }
    $allowed = @('schemaVersion', 'operation')
    if ($Request.operation -in @('folders', 'search')) { $allowed += @('mailboxes', 'folders', 'includeDescendants', 'allMailboxes', 'maxFolders', 'maxDurationSeconds') }
    if ($Request.operation -eq 'search') { $allowed += @('fromInclusive', 'toExclusive', 'allHistory', 'keywords', 'keywordGroups', 'listOnly', 'sender', 'summaryLimit', 'exhaustive', 'maxItems') }
    if ($Request.operation -eq 'browse') { $allowed += @('storeId', 'folderId', 'limit', 'maxItems', 'maxDurationSeconds') }
    if ($Request.operation -eq 'get') { $allowed += @('storeId', 'entryId', 'bodyOffset', 'bodyLength') }
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
    if ($Request.operation -in @('folders', 'search')) { Assert-MailLensQueryRequest $Request }
    if ($Request.operation -eq 'search') { Assert-MailLensSearchRequest $Request }
    if ($Request.operation -eq 'browse') { Assert-MailLensBrowseRequest $Request }
    if ($Request.operation -eq 'get') { Assert-MailLensDetailRequest $Request }
    if ('includeDescendants' -in $names -and $Request.includeDescendants -isnot [bool]) {
        Stop-MailLens 'invalid_request' 'includeDescendants must be a boolean.'
    }
}

function Add-MailLensProblem {
    param($Context, [string]$Code, [string]$Message, $Location = $null)
    $problem = [ordered]@{ code = $Code; message = $Message; location = $Location }
    if ($Context.ContainsKey('BoundedProblems') -and $Context.BoundedProblems) {
        $stage = $Context.Execution.stage; $property = $Context.Execution.property
        $problem.stage = $stage; $problem.property = $property; $problem.hResult = $Context.Execution.hResult
        $key = $Code + ':' + $stage + ':' + $property
        if (-not $Context.WarningGroups.ContainsKey($key)) {
            $Context.WarningGroups[$key] = [ordered]@{ code = $Code; stage = $stage; property = $property; count = 0; samples = 0 }
        }
        $group = $Context.WarningGroups[$key]; $group.count++; $Context.WarningCount++
        if ($Context.Problems.Count -ge 20 -or $group.samples -ge 2) { return }
        $group.samples++
    }
    [void]$Context.Problems.Add($problem)
}

function Assert-MailLensBrowseRequest {
    param($Request)
    $names = @($Request.PSObject.Properties.Name)
    if (('storeId' -in $names) -ne ('folderId' -in $names)) { Stop-MailLens 'invalid_request' 'storeId and folderId must be provided together.' }
    foreach ($field in @('storeId', 'folderId')) {
        if ($field -in $names -and ($Request.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($Request.$field))) {
            Stop-MailLens 'invalid_request' 'Browse identities must be nonempty strings.'
        }
    }
    $bounds = @{ limit = 100; maxItems = 10000; maxDurationSeconds = 300 }
    foreach ($field in $bounds.Keys) {
        if ($field -in $names -and ($Request.$field -isnot [int] -or $Request.$field -lt 1 -or $Request.$field -gt $bounds[$field])) {
            Stop-MailLens 'invalid_request' 'Browse limits must be integers within the documented ranges.'
        }
    }
    $limit = 10; $maxItems = 200
    if ('limit' -in $names) { $limit = $Request.limit }
    if ('maxItems' -in $names) { $maxItems = $Request.maxItems }
    if ($maxItems -lt $limit) { Stop-MailLens 'invalid_request' 'maxItems must be at least limit.' }
}

function Write-MailLensProgress {
    param($Execution, [string]$Stage, [int]$Count = 0)
    $elapsed = [double](& $Execution.Clock) - $Execution.started
    if ($Execution.stage -cne $Stage -or $elapsed - $Execution.lastProgress -ge 5) {
        $Execution.stage = $Stage; $Execution.lastProgress = $elapsed
        & $Execution.Progress ('MailLens phase={0} count={1} elapsedSeconds={2}' -f $Stage, $Count, [Math]::Max(0, [Math]::Floor($elapsed)))
    }
}

function Invoke-MailLensBoundedCall {
    param($Adapter, $Reference, [string]$Kind, [string]$Name, [object[]]$Arguments, $Execution, $CancellationRequested)
    Assert-MailLensNotCancelled $CancellationRequested
    if ([double](& $Execution.Clock) - $Execution.started -ge $Execution.duration) {
        Stop-MailLens 'time_budget' 'The cooperative time budget was reached; coverage is incomplete.'
    }
    Write-MailLensProgress $Execution $Execution.phase $Execution.visited
    $Execution.property = $Name
    $Execution.hResult = $null
    try {
        if ($Kind -eq 'Property') { return ,(& $Adapter.Property $Reference $Name) }
        return ,(& $Adapter.Method $Reference $Name $Arguments)
    } catch { $Execution.hResult = $_.Exception.HResult; throw }
}

function New-MailLensBoundedAdapter {
    param($Adapter, $Execution, $CancellationRequested)
    $original = $Adapter; $boundedCall = ${function:Invoke-MailLensBoundedCall}
    return @{
        Property = {
            param($Reference, $Name)
            & $boundedCall $original $Reference Property $Name @() $Execution $CancellationRequested
        }.GetNewClosure()
        Method = {
            param($Reference, $Name, $Arguments)
            & $boundedCall $original $Reference Method $Name $Arguments $Execution $CancellationRequested
        }.GetNewClosure()
        Release = $Adapter.Release
        Connect = $Adapter.Connect
    }
}

function New-MailLensDiagnosticAdapter {
    param($Adapter, $Execution, $CancellationRequested)
    $original = $Adapter
    return @{
        Property = {
            param($Reference, $Name)
            if (& $CancellationRequested) { $exception = [InvalidOperationException]::new('The operation was cancelled; requested coverage is incomplete.'); $exception.Data['MailLensCode']='cancelled'; throw $exception }
            $Execution.property = $Name; $Execution.hResult = $null
            try { return ,(& $original.Property $Reference $Name) } catch { $Execution.hResult = $_.Exception.HResult; throw }
        }.GetNewClosure()
        Method = {
            param($Reference, $Name, $Arguments)
            if (& $CancellationRequested) { $exception = [InvalidOperationException]::new('The operation was cancelled; requested coverage is incomplete.'); $exception.Data['MailLensCode']='cancelled'; throw $exception }
            $Execution.property = $Name; $Execution.hResult = $null
            try { return ,(& $original.Method $Reference $Name $Arguments) } catch { $Execution.hResult = $_.Exception.HResult; throw }
        }.GetNewClosure()
        Release = $Adapter.Release
        Connect = $Adapter.Connect
    }
}

function Read-MailLensDirectFolder {
    param($Adapter, $Session, $Request, $Context, $Response, $CancellationRequested)
    $execution = $Context.Execution; $store = $null; $folder = $null
    try {
        $execution.phase = 'acquire-folder'
        if ('storeId' -in @($Request.PSObject.Properties.Name)) {
            try { $folder = Invoke-MailLensBoundedCall $Adapter $Session Method GetFolderFromID @($Request.folderId, $Request.storeId) $execution $CancellationRequested }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                Stop-MailLens 'folder_unavailable' 'The exact requested folder is unavailable.' $_.Exception
            }
        } else {
            try { $store = Invoke-MailLensBoundedCall $Adapter $Session Property DefaultStore @() $execution $CancellationRequested }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                Stop-MailLens 'default_store_unavailable' 'The existing default store is unavailable.' $_.Exception
            }
            if ($null -eq $store) { Stop-MailLens 'default_store_unavailable' 'The existing default store is unavailable.' }
            try { $folder = Invoke-MailLensBoundedCall $Adapter $store Method GetDefaultFolder @(6) $execution $CancellationRequested }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                Stop-MailLens 'default_inbox_unavailable' 'The existing default Inbox is unavailable.' $_.Exception
            }
            if ($null -eq $folder) { Stop-MailLens 'default_inbox_unavailable' 'The existing default Inbox is unavailable.' }
        }
        if ($null -eq $folder) { Stop-MailLens 'folder_unavailable' 'The exact requested folder is unavailable.' }
        try {
            $storeId = [string](Invoke-MailLensBoundedCall $Adapter $folder Property StoreID @() $execution $CancellationRequested)
            $folderId = [string](Invoke-MailLensBoundedCall $Adapter $folder Property EntryID @() $execution $CancellationRequested)
            if ([string]::IsNullOrEmpty($storeId) -or [string]::IsNullOrEmpty($folderId)) { throw 'Missing identity.' }
            if ('storeId' -in @($Request.PSObject.Properties.Name) -and ($storeId -cne $Request.storeId -or $folderId -cne $Request.folderId)) { throw 'Changed identity.' }
        } catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            Stop-MailLens 'folder_identity_unavailable' 'The requested folder identity is unavailable or changed.' $_.Exception
        }
        $row = [ordered]@{ storeId = $storeId; storeName = $null; folderId = $folderId; folderPath = $null; readable = $true }
        $Response.scope = [ordered]@{ stores = @([ordered]@{ storeId = $storeId; storeName = $null; readable = $true }); folders = @($row) }
        try { $row.folderPath = [string](Invoke-MailLensBoundedCall $Adapter $folder Property FolderPath @() $execution $CancellationRequested) }
        catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            Add-MailLensProblem $Context 'display_metadata_unavailable' 'The folder display path is unavailable.'
        }
        # Store display metadata is optional; acquiring Store never discovers other stores.
        if ($null -eq $store) {
            try { $store = Invoke-MailLensBoundedCall $Adapter $folder Property Store @() $execution $CancellationRequested }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                Add-MailLensProblem $Context 'display_metadata_unavailable' 'The store display metadata is unavailable.'
            }
        }
        if ($null -ne $store) {
            try { $row.storeName = [string](Invoke-MailLensBoundedCall $Adapter $store Property DisplayName @() $execution $CancellationRequested) }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                Add-MailLensProblem $Context 'display_metadata_unavailable' 'The store display name is unavailable.'
            }
        }
        $Response.scope.stores[0].storeName = $row.storeName
        $result = $folder; $folder = $null
        return ,$result
    } finally {
        Close-MailLensReference $Adapter $folder $Context
        Close-MailLensReference $Adapter $store $Context
    }
}

function Read-MailLensBrowse {
    param($Adapter, $Session, $Request, $Context, $Response, $CancellationRequested)
    $data = $Response.data.browse; $execution = $Context.Execution
    $folder = $null; $items = $null; $summaries = New-Object 'System.Collections.Generic.List[object]'
    try {
        $folder = Read-MailLensDirectFolder $Adapter $Session $Request $Context $Response $CancellationRequested
        $row = $Response.scope.folders[0]; $execution.phase = 'acquire-items'
        try {
            $items = Invoke-MailLensBoundedCall $Adapter $folder Property Items @() $execution $CancellationRequested
            $count = [int](Invoke-MailLensBoundedCall $Adapter $items Property Count @() $execution $CancellationRequested)
            if ($null -eq $items -or $count -lt 0) { throw 'Unavailable Items.' }
        } catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            Stop-MailLens 'items_unavailable' 'The selected folder Items collection is unavailable.' $_.Exception
        }
        $execution.phase = 'sort-items'
        try { $null = Invoke-MailLensBoundedCall $Adapter $items Method Sort @('[ReceivedTime]', $true) $execution $CancellationRequested }
        catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            $data.ordering = 'collectionOrder'
            Add-MailLensProblem $Context 'sort_unavailable' 'Sorting is unavailable; this is a bounded collection-order sample, not the latest messages.'
        }
        $execution.phase = 'browse-items'
        for ($index = 1; $index -le $count; $index++) {
            if ($summaries.Count -ge $data.limit) { $data.stopReason = 'limit'; break }
            if ($data.visitedCount -ge $data.maxItems) { $data.stopReason = 'item_budget'; break }
            Assert-MailLensNotCancelled $CancellationRequested
            if ([double](& $execution.Clock) - $execution.started -ge $execution.duration) { Stop-MailLens 'time_budget' 'The cooperative time budget was reached; coverage is incomplete.' }
            $item = $null
            $data.visitedCount++; $execution.visited = $data.visitedCount
            try {
                $item = Invoke-MailLensBoundedCall $Adapter $items Method Item @($index) $execution $CancellationRequested
                if ([int](Invoke-MailLensBoundedCall $Adapter $item Property Class @() $execution $CancellationRequested) -ne 43) { $data.nonMailCount++; continue }
                $entryId = [string](Invoke-MailLensBoundedCall $Adapter $item Property EntryID @() $execution $CancellationRequested)
                $subject = Invoke-MailLensBoundedCall $Adapter $item Property Subject @() $execution $CancellationRequested
                if ([string]::IsNullOrEmpty($entryId) -or $subject -isnot [string]) { throw 'Missing usable metadata.' }
                $parent = $null
                try {
                    $parent = Invoke-MailLensBoundedCall $Adapter $item Property Parent @() $execution $CancellationRequested
                    $currentStore = [string](Invoke-MailLensBoundedCall $Adapter $parent Property StoreID @() $execution $CancellationRequested)
                    $currentFolder = [string](Invoke-MailLensBoundedCall $Adapter $parent Property EntryID @() $execution $CancellationRequested)
                    $currentItem = [string](Invoke-MailLensBoundedCall $Adapter $item Property EntryID @() $execution $CancellationRequested)
                    if ($currentStore -cne $row.storeId -or $currentFolder -cne $row.folderId -or $currentItem -cne $entryId) { throw 'Item moved or identity changed.' }
                } finally { Close-MailLensReference $Adapter $parent $Context }
                $summary = [ordered]@{ storeId = $row.storeId; storeName = $row.storeName; folderId = $row.folderId; folderPath = $row.folderPath
                    entryId = $entryId; subject = $subject; timestamp = $null; timestampField = $null; preview = $null; previewSource = $null }
                [void]$summaries.Add($summary)
                try {
                    $time = Convert-MailLensMailTime (Invoke-MailLensBoundedCall $Adapter $item Property ReceivedTime @() $execution $CancellationRequested) ([TimeZoneInfo]::Local)
                    if ($null -eq $time) { throw 'Unavailable date.' }
                    $summary.timestamp = $time.ToString('o'); $summary.timestampField = 'ReceivedTime'
                } catch {
                    if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                    Add-MailLensProblem $Context 'timestamp_unavailable' 'The optional ReceivedTime display timestamp is unavailable.'
                }
            } catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                $data.unreadableCount++
                Add-MailLensProblem $Context 'browse_item_unavailable' 'An attempted item has unavailable identity or subject metadata.'
            } finally { Close-MailLensReference $Adapter $item $Context }
        }
        $data.scanComplete = ($data.visitedCount -eq $count)
        $data.moreAvailable = (-not $data.scanComplete)
        if ($data.scanComplete) { $data.stopReason = 'end' }
        elseif ($null -eq $data.stopReason) { $data.stopReason = 'limit' }
        if ($data.stopReason -eq 'item_budget') { Add-MailLensProblem $Context 'item_budget' 'The item budget was reached; collection coverage is incomplete.' }
    } catch {
        $code = $_.Exception.Data['MailLensCode']
        if ($code -in @('cancelled', 'time_budget')) {
            $data.stopReason = $code; $data.scanComplete = $false
            if (Get-Variable count -Scope Local -ErrorAction SilentlyContinue) { $data.moreAvailable = ($data.visitedCount -lt $count) }
            if ($code -eq 'time_budget') { Add-MailLensProblem $Context 'time_budget' 'The cooperative time budget was reached; collection coverage is incomplete.'; return }
        }
        throw
    } finally {
        $data.summaries = @($summaries.ToArray()); $data.returnedCount = $summaries.Count
        Close-MailLensReference $Adapter $items $Context
        Close-MailLensReference $Adapter $folder $Context
    }
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
    if (& $CancellationRequested) { Stop-MailLens 'cancelled' 'The operation was cancelled; requested coverage is incomplete.' }
}

function Assert-MailLensQueryRequest {
    param($Request)
    $names = @($Request.PSObject.Properties.Name)
    foreach ($field in @('allMailboxes', 'exhaustive')) {
        if ($field -in $names -and $Request.$field -isnot [bool]) { Stop-MailLens 'invalid_request' 'Query mode flags must be booleans.' }
    }
    if ('allMailboxes' -in $names -and $Request.allMailboxes -and ('mailboxes' -in $names -or 'folders' -in $names)) {
        Stop-MailLens 'invalid_request' 'allMailboxes true cannot be combined with narrower selections.'
    }
    $bounds = @{ maxItems = 100000; maxFolders = 1000; maxDurationSeconds = 300 }
    foreach ($field in $bounds.Keys) {
        if ($field -in $names -and ($Request.$field -isnot [int] -or $Request.$field -lt 1 -or $Request.$field -gt $bounds[$field])) {
            Stop-MailLens 'invalid_request' 'Query budgets must be integers within the documented ranges.'
        }
    }
}

function Get-MailLensFolderKey {
    param([string]$StoreId, [string]$FolderId)
    return $StoreId.Length.ToString() + ':' + $StoreId + $FolderId
}

function Assert-MailLensFolderBudget {
    param($Context)
    if ($Context.FolderAttempts -ge $Context.Query.maxFolders) { Stop-MailLens 'folder_budget' 'The global folder budget was reached; requested scope is incomplete.' }
}

function Read-MailLensQueryFolderRow {
    param($Adapter, $Folder, $StoreRow, $ParentId, $Context)
    $Context.Execution.phase = 'resolve-folder'
    $row = [ordered]@{ storeId = $StoreRow.storeId; storeName = $StoreRow.storeName; folderId = $null
        folderPath = $null; parentFolderId = $ParentId; defaultItemType = $null; readable = $false; childrenReadable = $null }
    try {
        $row.folderId = [string](& $Adapter.Property $Folder 'EntryID')
        if ([string]::IsNullOrEmpty($row.folderId) -or [string](& $Adapter.Property $Folder 'StoreID') -cne $StoreRow.storeId) { throw 'Unavailable identity.' }
    } catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
        Assert-MailLensFolderBudget $Context; $Context.FolderAttempts++
        [void]$Context.Folders.Add($row)
        Add-MailLensProblem $Context 'folder_identity_unavailable' 'An attempted folder identity is unavailable.' $row
        return $row
    }
    $key = Get-MailLensFolderKey $row.storeId $row.folderId
    if (-not $Context.Seen.Contains($key)) {
        Assert-MailLensFolderBudget $Context; $Context.FolderAttempts++
        [void]$Context.Seen.Add($key); [void]$Context.Folders.Add($row)
    } else {
        return @($Context.Folders.ToArray() | Where-Object { $_.storeId -ceq $row.storeId -and $_.folderId -ceq $row.folderId })[0]
    }
    $row.readable = $true
    foreach ($property in @('FolderPath', 'DefaultItemType')) {
        try {
            $value = & $Adapter.Property $Folder $property
            if ($property -eq 'FolderPath') { $row.folderPath = [string]$value } else { $row.defaultItemType = [int]$value }
        } catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            Add-MailLensProblem $Context 'display_metadata_unavailable' 'A folder display property is unavailable.' $row
        }
    }
    return $row
}

function Read-MailLensQueryStoreRow {
    param($Adapter, $Store, $Context)
    $row = [ordered]@{ storeId = $null; storeName = $null; readable = $false }
    $row.storeId = [string](& $Adapter.Property $Store 'StoreID')
    if ([string]::IsNullOrEmpty($row.storeId)) { Stop-MailLens 'store_identity_unavailable' 'The selected store identity is unavailable.' }
    try { $row.storeName = [string](& $Adapter.Property $Store 'DisplayName') }
    catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
        Add-MailLensProblem $Context 'display_metadata_unavailable' 'The selected store display name is unavailable.'
    }
    $row.readable = $true
    $existing = @($Context.Stores.ToArray() | Where-Object { $_.storeId -ceq $row.storeId })
    if ($existing.Count -gt 0) { return $existing[0] }
    [void]$Context.Stores.Add($row)
    return $row
}

function Visit-MailLensQueryFolder {
    param($Adapter, $Folder, $StoreRow, $ParentId, [int]$Depth, $Context, $Search, $Paths, $CancellationRequested)
    $row = Read-MailLensQueryFolderRow $Adapter $Folder $StoreRow $ParentId $Context
    if (-not $row.readable) { return }
    $key = Get-MailLensFolderKey $row.storeId $row.folderId
    if (-not $Context.WalkSeen.Add($key)) { return }
    if ($null -eq $Paths) { $Context.SelectionResolved = $true }
    if ($null -ne $Search) { Read-MailLensSearchFolder $Adapter $Folder $row $Context $Search $CancellationRequested }
    if ($null -ne $Paths) {
        foreach ($selector in $Paths.Keys) {
            if ([string]::Equals($row.folderPath, $selector, [StringComparison]::Ordinal)) { [void]$Paths[$selector].Add($row) }
        }
    }
    if ($Depth -eq 0) { return }
    $children = $null; $Context.Execution.phase = 'discover-folders'
    try {
        $children = & $Adapter.Property $Folder 'Folders'
        $count = [int](& $Adapter.Property $children 'Count')
        if ($count -lt 0) { throw 'Invalid collection.' }
        $row.childrenReadable = $true
        $knownChildren = ($Context.KnownChildren.ContainsKey($key) -and $Context.KnownChildren[$key] -eq $count)
        for ($index = 1; $index -le $count; $index++) {
            if (-not $knownChildren) { Assert-MailLensFolderBudget $Context }
            $child = $null
            try {
                $child = & $Adapter.Method $children 'Item' @($index)
                Visit-MailLensQueryFolder $Adapter $child $StoreRow $row.folderId ($Depth - 1) $Context $Search $Paths $CancellationRequested
            } catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit')) { throw }
                $row.childrenReadable = $false; $Context.FolderAttempts++
                Add-MailLensProblem $Context 'folder_denied' 'An attempted child folder is unavailable.' ([ordered]@{ storeId = $row.storeId; parentFolderId = $row.folderId; childIndex = $index })
            } finally { Close-MailLensReference $Adapter $child $Context }
        }
        if ($row.childrenReadable) { $Context.KnownChildren[$key] = $count }
    } catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit')) { throw }
        $row.childrenReadable = $false
        Add-MailLensProblem $Context 'folder_denied' 'The selected child-folder collection is unavailable.' $row
    } finally { Close-MailLensReference $Adapter $children $Context }
}

function Read-MailLensQuery {
    param($Adapter, $Session, $Request, $Context, $Response, $Search, $CancellationRequested)
    $names = @($Request.PSObject.Properties.Name)
    $all = ('allMailboxes' -in $names -and $Request.allMailboxes)
    $explicit = ($all -or 'mailboxes' -in $names -or 'folders' -in $names)
    $descendants = ($Request.operation -eq 'search' -and $explicit)
    if ('includeDescendants' -in $names) { $descendants = $Request.includeDescendants }
    $Response.scope.requested = [ordered]@{ allMailboxes = $all; mailboxes = @(); folders = @(); includeDescendants = $descendants; defaultInbox = ($Request.operation -eq 'search' -and -not $explicit) }
    foreach ($field in @('mailboxes', 'folders')) { if ($field -in $names) { $Response.scope.requested[$field] = @($Request.$field) } }
    $targets = New-Object 'System.Collections.Generic.List[object]'
    $selectedStores = New-Object 'System.Collections.Generic.List[object]'
    $stores = $null; $folder = $null; $store = $null
    try {
        $Context.Execution.phase = 'resolve-scope'
        # Exact pairs without mailbox names need no mounted-store enumeration.
        $pairsOnly = ('folders' -in $names -and 'mailboxes' -notin $names -and @($Request.folders | Where-Object { $_ -is [string] }).Count -eq 0)
        if ($pairsOnly) {
            foreach ($selector in $Request.folders) { [void]$targets.Add(@{ kind = 'identity'; selector = $selector; store = $null }) }
        } elseif ($all -or 'mailboxes' -in $names) {
            $stores = & $Adapter.Property $Session 'Stores'
            $count = [int](& $Adapter.Property $stores 'Count'); $Context.StoreCount = $count
            $metadata = New-Object 'System.Collections.Generic.List[object]'; $resolutionGap = $false
            for ($index = 1; $index -le $count; $index++) {
                $store = $null
                try {
                    $store = & $Adapter.Method $stores 'Item' @($index)
                    $id = [string](& $Adapter.Property $store 'StoreID'); $display = $null
                    if ([string]::IsNullOrEmpty($id)) { throw 'Missing store identity.' }
                    if ('mailboxes' -in $names -and $id -cin $Request.mailboxes) { $needName = $false } else { $needName = $true }
                    try { $display = [string](& $Adapter.Property $store 'DisplayName') }
                    catch { if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }; if ($needName) { throw } }
                    [void]$metadata.Add(@{ id = $id; name = $display; store = $store }); $store = $null
                } catch {
                    if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                    $resolutionGap = $true
                    Add-MailLensProblem $Context 'store_metadata_unavailable' 'Mounted-store metadata is incomplete; unresolved selectors cannot be widened.'
                } finally { Close-MailLensReference $Adapter $store $Context }
            }
            if ($all) { foreach ($candidate in $metadata) { [void]$selectedStores.Add($candidate) } }
            else {
                foreach ($selector in $Request.mailboxes) {
                    $matches = @($metadata | Where-Object { $_.id -ceq $selector })
                    if ($matches.Count -eq 0) {
                        if ($resolutionGap) { Stop-MailLens 'selection_incomplete' 'A mailbox name cannot be resolved reliably because store metadata is incomplete.' }
                        $matches = @($metadata | Where-Object { $_.name -ceq $selector })
                    }
                    if ($matches.Count -eq 0) { Stop-MailLens 'unknown_selection' 'The requested mailbox does not match a mounted store.' }
                    if ($matches.Count -ne 1) {
                        foreach ($choice in $matches) { $null = Read-MailLensQueryStoreRow $Adapter $choice.store $Context }
                        Stop-MailLens 'ambiguous_selection' 'The mailbox selector is ambiguous; use an exact storeId.' }
                    if ($selectedStores -notcontains $matches[0]) { [void]$selectedStores.Add($matches[0]) }
                }
            }
            foreach ($candidate in $selectedStores) { $null = Read-MailLensQueryStoreRow $Adapter $candidate.store $Context }
            if ('folders' -notin $names) { foreach ($candidate in $selectedStores) { [void]$targets.Add(@{ kind = 'root'; store = $candidate.store }) } }
        } else {
            try { $store = & $Adapter.Property $Session 'DefaultStore' }
            catch { if ($_.Exception.Data['MailLensCode'] -in @('cancelled','time_budget')) { throw }; Stop-MailLens 'default_store_unavailable' 'The existing default store is unavailable.' $_.Exception }
            if ($null -eq $store) { Stop-MailLens 'default_store_unavailable' 'The existing default store is unavailable.' }
            $null = Read-MailLensQueryStoreRow $Adapter $store $Context
            [void]$selectedStores.Add(@{ id = [string](& $Adapter.Property $store 'StoreID'); name = $null; store = $store }); $store = $null
            if ('folders' -notin $names) { [void]$targets.Add(@{ kind = $(if ($Request.operation -eq 'search') { 'inbox' } else { 'root' }); store = $selectedStores[0].store }) }
        }
        if ('folders' -in $names -and -not $pairsOnly) {
            foreach ($selector in $Request.folders) {
                if ($selector -isnot [string]) {
                    if ($selector.storeId -cnotin @($selectedStores | ForEach-Object { $_.id })) { Stop-MailLens 'unknown_selection' 'The folder identity is outside the selected stores.' }
                    [void]$targets.Add(@{ kind = 'identity'; selector = $selector; store = $null })
                }
            }
            $paths = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
            foreach ($selector in $Request.folders) { if ($selector -is [string] -and -not $paths.ContainsKey($selector)) { $paths[$selector] = New-Object 'System.Collections.Generic.List[object]' } }
            if ($paths.Count -gt 0) {
                $Context.WalkSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
                foreach ($candidate in $selectedStores) {
                    $folder = $null
                    try {
                        Assert-MailLensFolderBudget $Context
                        $folder = & $Adapter.Method $candidate.store 'GetRootFolder' @()
                        $storeRow = Read-MailLensQueryStoreRow $Adapter $candidate.store $Context
                        Visit-MailLensQueryFolder $Adapter $folder $storeRow $null -1 $Context $null $paths $CancellationRequested
                    } finally { Close-MailLensReference $Adapter $folder $Context; $folder = $null }
                }
                foreach ($selector in $paths.Keys) {
                    $matches = $paths[$selector]
                    if ($Context.WarningCount -gt 0) { Stop-MailLens 'selection_incomplete' 'The requested paths cannot be resolved reliably because directory discovery is incomplete.' }
                    if ($matches.Count -eq 0) { Stop-MailLens 'unknown_selection' 'The requested path was not found within the selected stores.' }
                    if ($matches.Count -ne 1) { Stop-MailLens 'ambiguous_selection' 'The requested path is ambiguous; use an exact storeId/folderId pair.' }
                    [void]$targets.Add(@{ kind = 'identity'; selector = $matches[0]; store = $null })
                }
            }
        }
        $Context.WalkSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        foreach ($target in $targets) {
            $folder = $null; $store = $null
            try {
                if ($target.kind -eq 'identity') {
                    $key = Get-MailLensFolderKey $target.selector.storeId $target.selector.folderId
                    if ($Context.WalkSeen.Contains($key)) { continue }
                    if (-not $Context.Seen.Contains($key)) { Assert-MailLensFolderBudget $Context }
                    try { $folder = & $Adapter.Method $Session 'GetFolderFromID' @($target.selector.folderId, $target.selector.storeId) }
                    catch { if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }; Stop-MailLens 'folder_unavailable' 'The exact requested folder is unavailable.' $_.Exception }
                    if ([string](& $Adapter.Property $folder 'StoreID') -cne $target.selector.storeId -or [string](& $Adapter.Property $folder 'EntryID') -cne $target.selector.folderId) { Stop-MailLens 'folder_identity_changed' 'The exact requested folder identity changed.' }
                    try { $store = & $Adapter.Property $folder 'Store' }
                    catch {
                        if ($_.Exception.Data['MailLensCode'] -in @('cancelled','time_budget')) { throw }
                        Add-MailLensProblem $Context 'display_metadata_unavailable' 'The selected folder store metadata is unavailable.'
                        $Context.SentCache[$target.selector.storeId] = @{known=$false;folderId=$null}
                    }
                } else {
                    Assert-MailLensFolderBudget $Context
                    $store = $target.store
                    # The target list owns this acquisition; do not release it here.
                    if ($target.kind -eq 'inbox') {
                        try { $folder = & $Adapter.Method $store 'GetDefaultFolder' @(6) }
                        catch { if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }; Stop-MailLens 'default_inbox_unavailable' 'The existing default Inbox is unavailable.' $_.Exception }
                        if ($null -eq $folder) { Stop-MailLens 'default_inbox_unavailable' 'The existing default Inbox is unavailable.' }
                    } else { $folder = & $Adapter.Method $store 'GetRootFolder' @() }
                }
                if ($null -ne $store) { $storeRow = Read-MailLensQueryStoreRow $Adapter $store $Context }
                else {
                    $storeRow = [ordered]@{storeId=$target.selector.storeId;storeName=$null;readable=$true}
                    if (@($Context.Stores.ToArray() | Where-Object {$_.storeId -ceq $storeRow.storeId}).Count -eq 0) { [void]$Context.Stores.Add($storeRow) }
                }
                if ($target.kind -eq 'inbox') { $Context.KnownInboxKey = Get-MailLensFolderKey $storeRow.storeId ([string](& $Adapter.Property $folder 'EntryID')) }
                $depth = 0
                if ($descendants) { $depth = -1 } elseif ($Request.operation -eq 'folders' -and $target.kind -eq 'root') { $depth = 1 }
                Visit-MailLensQueryFolder $Adapter $folder $storeRow $null $depth $Context $Search $null $CancellationRequested
            } catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit', 'folder_unavailable', 'default_inbox_unavailable', 'folder_identity_changed')) { throw }
                if ($target.kind -eq 'root') {
                    $deniedId = [string](& $Adapter.Property $target.store 'StoreID')
                    foreach ($known in $Context.Stores) { if ($known.storeId -ceq $deniedId) { $known.readable = $false } }
                    Assert-MailLensFolderBudget $Context; $Context.FolderAttempts++
                    [void]$Context.Folders.Add([ordered]@{ storeId=$deniedId; folderId=$null; folderPath=$null; readable=$false; childrenReadable=$false })
                if ($null -ne $Search) { [void]$Search.data.folderCoverage.Add([ordered]@{storeId=$deniedId;folderId=$null;folderPath=$null;visitedCount=0;mailItemCount=0;inRangeCount=0;matchedCount=0;unreadableCount=1;complete=$false;scanComplete=$false;remainingCount=$null;dateFilter='notAttempted';timestampClassificationComplete=$false}) }
                }
                Add-MailLensProblem $Context 'store_denied' 'An attempted selected store or root is unavailable.'
            } finally {
                Close-MailLensReference $Adapter $folder $Context; $folder = $null
                if ($target.kind -eq 'identity') { Close-MailLensReference $Adapter $store $Context }; $store = $null
            }
        }
        if ($Context.Folders.Count -eq 0 -or @($Context.Folders.ToArray() | Where-Object { $_.readable }).Count -eq 0) { Stop-MailLens 'selection_unavailable' 'No trustworthy folder target could be resolved.' }
        $Context.Query.scanComplete = (@($Context.Folders.ToArray() | Where-Object { -not $_.readable -or $_.childrenReadable -eq $false }).Count -eq 0); $Context.Query.stopReason = 'end'
        if ($null -ne $Search) { $Search.data.scanComplete = ($Context.Query.scanComplete -and @($Search.data.folderCoverage.ToArray() | Where-Object { -not $_.scanComplete }).Count -eq 0); $Search.data.stopReason = 'end' }
    } finally {
        Close-MailLensReference $Adapter $folder $Context
        Close-MailLensReference $Adapter $store $Context
        Close-MailLensReference $Adapter $stores $Context
        if (Get-Variable metadata -Scope Local -ErrorAction SilentlyContinue) { foreach ($candidate in $metadata) { Close-MailLensReference $Adapter $candidate.store $Context } }
        else { foreach ($candidate in $selectedStores) { Close-MailLensReference $Adapter $candidate.store $Context } }
    }
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
    if ('summaryLimit' -in $names -and ($Request.summaryLimit -isnot [int] -or $Request.summaryLimit -lt 1 -or $Request.summaryLimit -gt 100)) {
        Stop-MailLens 'invalid_request' 'summaryLimit must be an integer from 1 to 100.'
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
        summaryLimit = 20; exhaustive = $false; maxItems = 1000; maxFolders = 25; maxDurationSeconds = 30
        visitedCount = 0; attemptedFolderCount = 0; scanComplete = $false; matchedCountIsExact = $false; stopReason = $null; ordering = 'retainedEffectiveTimeDescending'; matchedCount = 0; returnedCount = 0; resultsTruncated = $false
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
    foreach ($field in @('summaryLimit', 'exhaustive', 'maxItems', 'maxFolders', 'maxDurationSeconds')) { if ($field -in $names) { $data[$field] = $Request.$field } }
    return @{ data = $data; zone = $zone; from = $from; to = $to; matches = New-Object 'System.Collections.Generic.List[object]' }
}

function Read-MailLensSentClassification {
    param($Adapter, $Folder, $Row, $Context)
    $Context.Execution.phase = 'classify-timestamp'
    if ((Get-MailLensFolderKey $Row.storeId $Row.folderId) -ceq $Context.KnownInboxKey) { return @{known=$true;isSent=$false} }
    if (-not $Context.SentCache.ContainsKey($Row.storeId)) {
        $store = $null; $sent = $null
        $classification = @{ known = $false; folderId = $null }
        try {
            $store = & $Adapter.Property $Folder 'Store'
            $sent = & $Adapter.Method $store 'GetDefaultFolder' @(5)
            # A store with no Sent folder is a known non-Sent tree.
            if ($null -ne $sent) {
                $classification.folderId = [string](& $Adapter.Property $sent 'EntryID')
                if ([string]::IsNullOrEmpty($classification.folderId)) { throw 'Missing Sent identity.' }
            }
            $classification.known = $true
        } catch {
            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
            Add-MailLensProblem $Context 'sent_folder_denied' 'Sent Items classification is unavailable; fallback timestamps cannot certify exact time coverage.' $Row
        } finally { Close-MailLensReference $Adapter $sent $Context; Close-MailLensReference $Adapter $store $Context }
        $Context.SentCache[$Row.storeId] = $classification
    }
    $classification = $Context.SentCache[$Row.storeId]
    if (-not $classification.known) { return @{ known = $false; isSent = $false } }
    if ($null -eq $classification.folderId) { return @{ known = $true; isSent = $false } }
    # Already discovered ancestors use identities, never subject/path substitutions.
    $ancestor = $Row
    while ($null -ne $ancestor) {
        if ($ancestor.folderId -ceq $classification.folderId) { return @{ known = $true; isSent = $true } }
        if ($null -eq $ancestor.parentFolderId) { break }
        $matches = @($Context.Folders.ToArray() | Where-Object { $_.storeId -ceq $Row.storeId -and $_.folderId -ceq $ancestor.parentFolderId })
        if ($matches.Count -eq 0) { break }; $ancestor = $matches[0]
    }
    if ($null -ne $Row.parentFolderId) { return @{ known = $true; isSent = $false } }
    # Directly selected nested folders can predate discovery of their ancestry.
    $current = $null; $parent = $null
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    [void]$seen.Add($Row.folderId)
    try {
        $current = & $Adapter.Property $Folder 'Parent'
        while ($null -ne $current) {
            # Namespace parents are not folders and need no further traversal.
            try { $id = [string](& $Adapter.Property $current 'EntryID') }
            catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                if ([int](& $Adapter.Property $current 'Class') -eq 1) { break }
                throw
            }
            if ([string](& $Adapter.Property $current 'StoreID') -cne $Row.storeId -or -not $seen.Add($id)) { throw 'Unreliable ancestry.' }
            $key = Get-MailLensFolderKey $Row.storeId $id
            if (-not $Context.Seen.Contains($key)) {
                Assert-MailLensFolderBudget $Context
                $null = Read-MailLensQueryFolderRow $Adapter $current ([ordered]@{storeId=$Row.storeId;storeName=$Row.storeName}) $null $Context
            }
            if ($id -ceq $classification.folderId) { return @{ known = $true; isSent = $true } }
            $parent = & $Adapter.Property $current 'Parent'
            Close-MailLensReference $Adapter $current $Context; $current = $parent; $parent = $null
        }
        return @{ known = $true; isSent = $false }
    } catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget')) { throw }
        Add-MailLensProblem $Context 'sent_ancestry_unavailable' 'Selected-folder ancestry is unavailable; fallback timestamps are qualified.' $Row
        return @{ known = $false; isSent = $false }
    } finally { Close-MailLensReference $Adapter $parent $Context; Close-MailLensReference $Adapter $current $Context }
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
    elseif ($SentKnown -and $IsSentFolder) { $fields = @('SentOn') }
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

function New-MailLensDateFilter {
    param($Search)
    if ($null -eq $Search.from) { return $null }
    # DASL date comparisons use UTC and Windows regional date/time strings.
    # Widen both boundaries to whole minutes: exact half-open checks follow.
    $fromTicks = $Search.from.UtcDateTime.Ticks
    $minute = [TimeSpan]::TicksPerMinute
    $from = [DateTime]::new(($fromTicks - ($fromTicks % $minute)), [DateTimeKind]::Utc)
    $toTicks = $Search.to.UtcDateTime.Ticks
    $to = [DateTime]::new(($toTicks - ($toTicks % $minute)), [DateTimeKind]::Utc)
    if (($toTicks % $minute) -ne 0) { $to = $to.AddMinutes(1) }
    $culture = [Globalization.CultureInfo]::CurrentCulture
    $pattern = $culture.DateTimeFormat.ShortDatePattern + ' ' + $culture.DateTimeFormat.ShortTimePattern
    $lower = $from.ToString($pattern, $culture).Replace("'", "''")
    $upper = $to.ToString($pattern, $culture).Replace("'", "''")
    # Preserve null/unset candidate fields as well as in-range dates. Any
    # candidate field can be the effective date for drafts, Sent mail or the fallback.
    # A day of padding also covers local/UTC offsets at the invalid boundaries.
    $invalidLower = ([DateTime]::new(1601, 1, 2)).ToString($pattern, $culture).Replace("'", "''")
    $invalidUpper = ([DateTime]::new(4500, 12, 31)).ToString($pattern, $culture).Replace("'", "''")
    $parts = foreach ($property in @('urn:schemas:httpmail:datereceived', 'urn:schemas:httpmail:date', 'http://schemas.microsoft.com/mapi/proptag/0x30070040')) {
        '("' + $property + '" >= ''' + $lower + ''' AND "' + $property + '" < ''' + $upper + ''') OR ' +
            '("' + $property + '" IS NULL OR "' + $property + '" <= ''' + $invalidLower + ''' OR "' + $property + '" >= ''' + $invalidUpper + ''')'
    }
    return '@SQL=' + ($parts -join ' OR ')
}

function Read-MailLensOptionalModification {
    param($Adapter, $Item, $Context, $Location)
    try {
        $value = & $Adapter.Property $Item 'LastModificationTime'
        if ($value -isnot [DateTime] -or $value -le [DateTime]'1601-01-01' -or $value.Year -ge 4501) { throw 'Unavailable modification marker.' }
        return $value
    } catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
        Add-MailLensProblem $Context 'consistency_unavailable' 'LastModificationTime is unavailable; readable data is not a certified consistent snapshot.' $Location
        return $null
    }
}

function Read-MailLensSearchFolder {
    param($Adapter, $Folder, $Row, $Context, $Search, $CancellationRequested)
    $key = Get-MailLensFolderKey $Row.storeId $Row.folderId
    if (-not $Context.ScanSeen.Add($key)) { return }
    $coverage = [ordered]@{ storeId = $Row.storeId; folderId = $Row.folderId; folderPath = $Row.folderPath
        visitedCount = 0; mailItemCount = 0; inRangeCount = 0; matchedCount = 0; unreadableCount = 0
        complete = $false; scanComplete = $false; remainingCount = $null; dateFilter = 'notRequested'; timestampClassificationComplete = $true }
    [void]$Search.data.folderCoverage.Add($coverage)
    $items = $null; $restricted = $null
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $baselineWarnings = $Context.WarningCount
    try {
        $classification = Read-MailLensSentClassification $Adapter $Folder $Row $Context
        $coverage.timestampClassificationComplete = $classification.known
        $Context.Execution.phase = 'acquire-items'
        $items = & $Adapter.Property $Folder 'Items'
        if ($null -eq $items) { throw 'Unavailable Items.' }
        if ($null -ne $Search.from) {
            $Context.Execution.phase = 'restrict-date'
            try {
                $filter = New-MailLensDateFilter $Search
                $restricted = & $Adapter.Method $items 'Restrict' @($filter)
                if ($null -eq $restricted) { throw 'Unavailable restricted collection.' }
                $coverage.dateFilter = 'coarseUtcUnion'
            } catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                $coverage.dateFilter = 'boundedFallback'
                Add-MailLensProblem $Context 'date_filter_unavailable' 'Date Restrict is unavailable; this folder uses the original collection under the global budgets.' $Row
            }
        }
        $candidates = $items; if ($null -ne $restricted) { $candidates = $restricted }
        $count = [int](& $Adapter.Property $candidates 'Count')
        if ($count -lt 0) { throw 'Invalid item count.' }
        $Context.Execution.phase = 'search-items'
        $deniedAttempts = 0
        for ($index = 1; $index -le $count; $index++) {
            if ($Search.data.visitedCount -ge $Search.data.maxItems) { Stop-MailLens 'item_budget' 'The global item budget was reached.' }
            $item = $null
            $location = [ordered]@{ storeId = $Row.storeId; folderId = $Row.folderId; itemIndex = $index; entryId = $null }
            $coverage.visitedCount++; $Search.data.visitedCount++; $Context.Execution.visited = $Search.data.visitedCount
            try {
                $item = & $Adapter.Method $candidates 'Item' @($index)
                if ([int](& $Adapter.Property $item 'Class') -ne 43) { continue }
                $coverage.mailItemCount++
                $entryId = [string](& $Adapter.Property $item 'EntryID'); $location.entryId = $entryId
                if ([string]::IsNullOrEmpty($entryId) -or -not $seen.Add($entryId)) { throw 'Missing/repeated identity.' }
                Assert-MailLensItemLocation $Adapter $item $Row $Context
                $time = Read-MailLensEffectiveTime $Adapter $item $classification.isSent $classification.known $Search.zone
                if ($null -ne $Search.from -and ($time.value -lt $Search.from -or $time.value -ge $Search.to)) { continue }
                $coverage.inRangeCount++
                $sender = $null
                if ($null -ne $Search.data.sender) {
                    $sender = [ordered]@{ displayName = $null; rawAddress = $null; addressType = $null }
                    $available = 0; $senderGap = $false
                    foreach ($pair in @(@('SenderName', 'displayName'), @('SenderEmailAddress', 'rawAddress'), @('SenderEmailType', 'addressType'))) {
                        try {
                            $value = & $Adapter.Property $item $pair[0]
                            if ($value -isnot [string]) { throw 'Unavailable sender.' }
                            $sender[$pair[1]] = $value; if ($pair[1] -ne 'addressType') { $available++ }
                        } catch {
                            if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }
                            $senderGap = $true
                            Add-MailLensProblem $Context 'sender_denied' 'Sender metadata is partly unavailable; narrowing coverage is incomplete.' $location
                        }
                    }
                    if ($senderGap) { $coverage.unreadableCount++ }
                    if ($available -eq 0) { throw 'Unavailable sender identity.' }
                    if (([string]$sender.displayName).IndexOf($Search.data.sender, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and
                        ([string]$sender.rawAddress).IndexOf($Search.data.sender, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                }
                $modified = Read-MailLensOptionalModification $Adapter $item $Context $location
                $subject = & $Adapter.Property $item 'Subject'
                if ($subject -isnot [string]) { throw 'Unavailable subject.' }
                $body = $null
                if (-not (Test-MailLensTerms $subject '' $Search.data)) {
                    $body = & $Adapter.Property $item 'Body'
                    if ($body -isnot [string]) { throw 'Unavailable plain body.' }
                    if (-not (Test-MailLensTerms $subject $body $Search.data)) { continue }
                }
                Assert-MailLensItemLocation $Adapter $item $Row $Context
                if ([string](& $Adapter.Property $item 'EntryID') -cne $entryId) { throw 'Changed identity.' }
                $after = $null
                # A denied marker is never retried on the same item.
                if ($null -ne $modified) { $after = Read-MailLensOptionalModification $Adapter $item $Context $location }
                if ($null -ne $modified -and $null -ne $after -and $modified -ne $after) { throw 'Content changed during evaluation.' }
                $coverage.matchedCount++; $Search.data.matchedCount++
                if ($Search.matches.Count -lt $Search.data.summaryLimit) {
                    $preview = $null; $previewSource = $null
                    if ($null -ne $body) {
                        $length = [Math]::Min(500, $body.Length)
                        if ($length -gt 0 -and [char]::IsHighSurrogate($body[$length - 1])) { $length-- }
                        $preview = $body.Substring(0, $length); $previewSource = 'Body'
                    }
                    [void]$Search.matches.Add([ordered]@{ storeId = $Row.storeId; storeName = $Row.storeName; folderId = $Row.folderId; folderPath = $Row.folderPath
                        entryId = $entryId; subject = $subject; timestamp = $time.value.ToUniversalTime().ToString('o'); timestampField = $time.field
                        timestampClassificationComplete = $classification.known; consistencyAvailable = ($null -ne $modified -and $null -ne $after)
                        preview = $preview; previewSource = $previewSource; sender = $sender })
                }
                if (-not $Search.data.exhaustive -and $Search.matches.Count -ge $Search.data.summaryLimit) { Stop-MailLens 'summary_limit' 'The bounded summary limit was reached; remaining scope is unknown.' }
            } catch {
                if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit')) { throw }
                $coverage.unreadableCount++
                $code = $_.Exception.Data['MailLensCode']; if ($null -eq $code) { $code = 'item_unavailable' }
                Add-MailLensProblem $Context $code 'An attempted item could not be evaluated reliably; matching coverage is incomplete.' $location
                # Repeated denied Item acquisition is a collection capability failure.
                # Body/property failures are individual item gaps, never retried.
                if ($null -eq $item -and $_.Exception.HResult -in @(-2147024891, -2147221246)) { $deniedAttempts++ }
                if ($deniedAttempts -ge 2) {
                    $coverage.remainingCount = $count - $coverage.visitedCount
                    Add-MailLensProblem $Context 'folder_capability_denied' 'Repeated collection access denial closed this folder path; remaining items are unknown.' $Row
                    return
                }
            } finally { Close-MailLensReference $Adapter $item $Context }
        }
        if ([int](& $Adapter.Property $candidates 'Count') -ne $count) { throw 'Collection changed.' }
        $coverage.scanComplete = $true; $coverage.remainingCount = 0
        $coverage.complete = ($coverage.unreadableCount -eq 0 -and $Context.WarningCount -eq $baselineWarnings)
    } catch {
        if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit')) { throw }
        $coverage.unreadableCount++
        Add-MailLensProblem $Context 'search_folder_denied' 'The selected item collection is unavailable or changed; matching coverage is incomplete.' $Row
    } finally {
        if (-not $coverage.scanComplete -and (Get-Variable count -Scope Local -ErrorAction SilentlyContinue)) { $coverage.remainingCount = [Math]::Max(0, $count - $coverage.visitedCount) }
        Close-MailLensReference $Adapter $restricted $Context
        Close-MailLensReference $Adapter $items $Context
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
    $Search.data.resultsTruncated = ($Search.data.matchedCount -gt $Search.data.returnedCount -or -not $Search.data.scanComplete)
    $Search.data.folderCoverage = @($Search.data.folderCoverage.ToArray())
}


function Assert-MailLensDetailRequest {
    param($Request)
    $names = @($Request.PSObject.Properties.Name)
    foreach ($field in @('storeId', 'entryId')) {
        if ($field -cnotin $names -or $Request.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($Request.$field)) {
            Stop-MailLens 'invalid_request' 'get requires nonempty string storeId and entryId.'
        }
    }
    if ('bodyOffset' -in $names -and ($Request.bodyOffset -isnot [int] -or $Request.bodyOffset -lt 0)) {
        Stop-MailLens 'invalid_request' 'bodyOffset must be a nonnegative integer.'
    }
    if ('bodyLength' -in $names -and ($Request.bodyLength -isnot [int] -or $Request.bodyLength -lt 1 -or $Request.bodyLength -gt 100000)) {
        Stop-MailLens 'invalid_request' 'bodyLength must be an integer from 1 to 100000.'
    }
}

function Read-MailLensDetailProperty {
    param($Adapter, $Reference, [string]$Name, $Context, $Location, $Zone = $null)
    try {
        $value = & $Adapter.Property $Reference $Name
        if ($null -ne $Zone) {
            $time = Convert-MailLensMailTime $value $Zone
            if ($null -eq $time) { throw 'Unavailable date.' }
            return $time.ToUniversalTime().ToString('o')
        }
        if ($value -isnot [string]) { throw 'Unavailable string.' }
        return $value
    } catch {
        if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
        Add-MailLensProblem $Context 'detail_property_unavailable' 'A requested detail property is unavailable.' ([ordered]@{ source = $Location; property = $Name })
        return $null
    }
}

function Read-MailLensDetailList {
    param($Adapter, $Item, [string]$Name, $Context, $Location, $CancellationRequested)
    $collection = $null
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $complete = $true
    try {
        Assert-MailLensNotCancelled $CancellationRequested
        $collection = & $Adapter.Property $Item $Name
        $count = [int](& $Adapter.Property $collection 'Count')
        if ($count -lt 0) { throw 'Invalid collection count.' }
        for ($index = 1; $index -le $count; $index++) {
            Assert-MailLensNotCancelled $CancellationRequested
            $child = $null; $addressEntry = $null
            $location = [ordered]@{ source = $Location; collection = $Name; index = $index }
            try {
                $child = & $Adapter.Method $collection 'Item' @($index)
                if ($Name -eq 'Attachments') {
                    $filename = Read-MailLensDetailProperty $Adapter $child 'FileName' $Context $location
                    if ($null -ne $filename) { [void]$rows.Add($filename) } else { $complete = $false }
                } else {
                    $row = [ordered]@{ displayName = $null; rawAddress = $null; addressType = $null; recipientType = $null }
                    $row.displayName = Read-MailLensDetailProperty $Adapter $child 'Name' $Context $location
                    $row.rawAddress = Read-MailLensDetailProperty $Adapter $child 'Address' $Context $location
                    try { $row.recipientType = [int](& $Adapter.Property $child 'Type') }
                    catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; $complete = $false; Add-MailLensProblem $Context 'recipient_unavailable' 'The recipient role is unavailable.' $location }
                    try {
                        # AddressEntry forces resolution if unresolved. Never cause that side effect.
                        if ((& $Adapter.Property $child 'Resolved') -ne $true) { throw 'Unresolved recipient.' }
                        $addressEntry = & $Adapter.Property $child 'AddressEntry'
                        $row.addressType = Read-MailLensDetailProperty $Adapter $addressEntry 'Type' $Context $location
                    } catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; $complete = $false; Add-MailLensProblem $Context 'recipient_unavailable' 'Recipient address type is unavailable; unresolved names are not resolved.' $location }
                    if ($null -eq $row.displayName -or $null -eq $row.rawAddress -or $null -eq $row.addressType) { $complete = $false }
                    [void]$rows.Add($row)
                }
            } catch {
                if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
                $complete = $false
                Add-MailLensProblem $Context 'detail_collection_unavailable' 'A requested detail collection entry is unavailable.' $location
            } finally {
                Close-MailLensReference $Adapter $addressEntry $Context
                Close-MailLensReference $Adapter $child $Context
            }
        }
        if ([int](& $Adapter.Property $collection 'Count') -ne $count) { throw 'Collection changed.' }
    } catch {
        if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }
        $complete = $false
        Add-MailLensProblem $Context 'detail_collection_unavailable' 'A requested detail collection is unavailable or changed.' ([ordered]@{ source = $Location; collection = $Name })
    } finally { Close-MailLensReference $Adapter $collection $Context }
    return @{ rows = @($rows.ToArray()); complete = $complete }
}

function Read-MailLensBodyChunk {
    param([string]$Body, [int]$Offset, [int]$Length, $CancellationRequested)
    # Request and response character offsets count Unicode scalar values, not UTF-16
    # code units or grapheme clusters. Keep surrogate pairs intact in each UTF-8 chunk.
    $total = 0; $start = $Body.Length; $end = $Body.Length
    for ($index = 0; $index -lt $Body.Length; $index++) {
        if (($total % 4096) -eq 0) { Assert-MailLensNotCancelled $CancellationRequested }
        if ($total -eq $Offset) { $start = $index }
        if ([long]$total -eq ([long]$Offset + $Length)) { $end = $index }
        if ([char]::IsHighSurrogate($Body[$index])) {
            if ($index + 1 -ge $Body.Length -or -not [char]::IsLowSurrogate($Body[$index + 1])) { Stop-MailLens 'body_unavailable' 'The plain body contains invalid Unicode.' }
            $index++
        } elseif ([char]::IsLowSurrogate($Body[$index])) { Stop-MailLens 'body_unavailable' 'The plain body contains invalid Unicode.' }
        $total++
    }
    $returned = 0
    if ($Offset -lt $total) { $returned = [Math]::Min($Length, $total - $Offset) }
    return [ordered]@{ body = $Body.Substring($start, $end - $start); bodySource = 'Body'
        bodyOffsetUnit = 'unicodeScalarValues'; bodyTotalLength = $total; bodyOffset = $Offset
        bodyLength = $Length; bodyReturnedLength = $returned; bodyTruncated = ([long]$Offset + $returned -lt $total) }
}

function Read-MailLensDetail {
    param($Adapter, $Session, $Request, $Context, $Response, $CancellationRequested)
    $item = $null; $parent = $null; $store = $null
    $source = [ordered]@{ storeId = $Request.storeId; storeName = $null; folderId = $null; folderPath = $null
        entryId = $Request.entryId; subject = $null; timestamp = $null; timestampField = 'LastModificationTime' }
    try {
        Assert-MailLensNotCancelled $CancellationRequested
        try { $item = & $Adapter.Method $Session 'GetItemFromID' @($Request.entryId, $Request.storeId) }
        catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; Stop-MailLens 'item_unavailable' 'The exact storeId/entryId item is stale, deleted, moved or inaccessible.' }
        try {
            if ([int](& $Adapter.Property $item 'Class') -ne 43) { Stop-MailLens 'unsupported_item_type' 'The requested item is not a MailItem.' }
            if ([string](& $Adapter.Property $item 'EntryID') -cne $Request.entryId) { throw 'Mismatched item.' }
            $parent = & $Adapter.Property $item 'Parent'
            if ([string](& $Adapter.Property $parent 'StoreID') -cne $Request.storeId) { throw 'Mismatched store.' }
            $source.folderId = [string](& $Adapter.Property $parent 'EntryID')
            if ([string]::IsNullOrEmpty($source.folderId)) { throw 'Missing folder identity.' }
        } catch {
            if ($_.Exception.Data['MailLensCode'] -in @('unsupported_item_type', 'cancelled')) { throw }
            Stop-MailLens 'item_identity_unavailable' 'The requested item type or exact source identity could not be verified.'
        }
        $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Response.data.runtime.timeZone)
        $modified = Read-MailLensDetailProperty $Adapter $item 'LastModificationTime' $Context $source $zone
        $source.timestamp = $modified
        Assert-MailLensNotCancelled $CancellationRequested
        try {
            $body = & $Adapter.Property $item 'Body'
            if ($body -isnot [string]) { throw 'Missing body.' }
        } catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; Stop-MailLens 'body_unavailable' 'The requested plain Body is unavailable; it is not an empty body.' }
        $offset = 0; $length = 2000
        $names = @($Request.PSObject.Properties.Name)
        if ('bodyOffset' -in $names) { $offset = $Request.bodyOffset }
        if ('bodyLength' -in $names) { $length = $Request.bodyLength }
        $detail = Read-MailLensBodyChunk $body $offset $length $CancellationRequested
        $detail.source = $source
        $detail.lastModificationTime = $modified; $detail.lastModificationTimeAfter = $null
        $detail.consistencyAvailable = ($null -ne $modified); $detail.contentChangedDuringRead = $false
        $source.subject = Read-MailLensDetailProperty $Adapter $item 'Subject' $Context $source
        $source.folderPath = Read-MailLensDetailProperty $Adapter $parent 'FolderPath' $Context $source
        try {
            $store = & $Adapter.Property $parent 'Store'
            if ($null -eq $store) { throw 'Unavailable store.' }
            if ([string](& $Adapter.Property $store 'StoreID') -cne $Request.storeId) { throw 'Changed store.' }
            $source.storeName = Read-MailLensDetailProperty $Adapter $store 'DisplayName' $Context $source
        } catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; Add-MailLensProblem $Context 'detail_property_unavailable' 'The current store display metadata is unavailable.' $source }
        $detail.sender = [ordered]@{ displayName = $null; rawAddress = $null; addressType = $null }
        foreach ($pair in @(@('SenderName', 'displayName'), @('SenderEmailAddress', 'rawAddress'), @('SenderEmailType', 'addressType'))) {
            Assert-MailLensNotCancelled $CancellationRequested
            $detail.sender[$pair[1]] = Read-MailLensDetailProperty $Adapter $item $pair[0] $Context $source
        }
        foreach ($pair in @(@('SentOn', 'sentOn'), @('ReceivedTime', 'receivedTime'), @('CreationTime', 'creationTime'))) {
            $detail[$pair[1]] = Read-MailLensDetailProperty $Adapter $item $pair[0] $Context $source $zone
        }
        $recipients = Read-MailLensDetailList $Adapter $item 'Recipients' $Context $source $CancellationRequested
        $attachments = Read-MailLensDetailList $Adapter $item 'Attachments' $Context $source $CancellationRequested
        $detail.recipients = $recipients.rows; $detail.recipientsComplete = $recipients.complete
        $detail.attachmentNames = $attachments.rows; $detail.attachmentNamesComplete = $attachments.complete
        Assert-MailLensNotCancelled $CancellationRequested
        try {
            if ([string](& $Adapter.Property $item 'EntryID') -cne $Request.entryId) { throw 'Changed item.' }
            Assert-MailLensItemLocation $Adapter $item $source $Context
        } catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; Stop-MailLens 'item_identity_changed' 'The item source changed during retrieval; reread the current identity.' }
        if ($null -ne $modified) { $detail.lastModificationTimeAfter = Read-MailLensDetailProperty $Adapter $item 'LastModificationTime' $Context $source $zone }
        if ($null -eq $detail.lastModificationTimeAfter) { $detail.consistencyAvailable = $false }
        if ($null -ne $modified -and $null -ne $detail.lastModificationTimeAfter -and $modified -cne $detail.lastModificationTimeAfter) {
            $detail.contentChangedDuringRead = $true; $detail.consistencyAvailable = $false
            Add-MailLensProblem $Context 'item_content_changed' 'The item changed during retrieval; discard combined chunks and reread.' $source
        }
        Assert-MailLensNotCancelled $CancellationRequested
        $Response.data.detail = $detail
        $Response.scope = [ordered]@{ stores = @([ordered]@{ storeId = $source.storeId; storeName = $source.storeName; readable = $true })
            folders = @([ordered]@{ storeId = $source.storeId; folderId = $source.folderId; folderPath = $source.folderPath; readable = $true }) }
    } finally {
        Close-MailLensReference $Adapter $store $Context
        Close-MailLensReference $Adapter $parent $Context
        Close-MailLensReference $Adapter $item $Context
    }
}

function Invoke-MailLensRequest {
    [CmdletBinding()]
    param($Request, $Adapter = (New-MailLensComAdapter), $Runtime = (Get-MailLensRuntime),
        [scriptblock]$CancellationRequested = { $false }, [DateTimeOffset]$StartedAt = [DateTimeOffset]::UtcNow,
        [scriptblock]$Clock = { [double][Diagnostics.Stopwatch]::GetTimestamp() / [Diagnostics.Stopwatch]::Frequency },
        [scriptblock]$Progress = { param($Line) [Console]::Error.WriteLine($Line) })
    $execution = @{ Clock = $Clock; Progress = $Progress; started = [double](& $Clock); lastProgress = -5; stage = ''; phase = ''; property = $null; visited = 0; duration = 300; hResult = $null }
    $context = @{
        Execution = $execution; BoundedProblems = $true; WarningGroups = @{}; WarningCount = 0
        Stores = New-Object 'System.Collections.Generic.List[object]'
        Folders = New-Object 'System.Collections.Generic.List[object]'
        Problems = New-Object 'System.Collections.Generic.List[object]'
        Seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        KnownChildren = @{}; KnownInboxKey = $null; SelectionResolved = $false; StoreCount = $null; FolderAttempts = 0; Query = $null; SentCache = @{}; WalkSeen = $null; ScanSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    }
    $response = [ordered]@{
        schemaVersion = 2; operation = $null; status = 'failed'
        startedAt = $StartedAt.ToString('o'); endedAt = $null
        scope = [ordered]@{ stores = @(); folders = @() }; warnings = @(); errors = @()
        data = [ordered]@{ runtime = $Runtime; sessionMode = $null; stores = @(); folders = @(); coverage = $null; discoveryAttempted = $false }
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
        if ($Request.operation -in @('folders', 'search')) {
            $query = [ordered]@{ maxFolders = 25; maxDurationSeconds = 30; attemptedFolderCount = 0; scanComplete = $false; stopReason = $null }
            foreach ($field in @('maxFolders', 'maxDurationSeconds')) { if ($field -in @($Request.PSObject.Properties.Name)) { $query[$field] = $Request.$field } }
            $context.Query = $query; $execution.duration = $query.maxDurationSeconds
            if ($Request.operation -eq 'folders') { $response.data.directory = $query }
        }
        if ($Request.operation -eq 'browse') {
            $context.BoundedProblems = $true
            $browse = [ordered]@{ limit = 10; maxItems = 200; maxDurationSeconds = 30; ordering = 'receivedTimeDescending'
                visitedCount = 0; returnedCount = 0; nonMailCount = 0; unreadableCount = 0; scanComplete = $false; moreAvailable = $null; stopReason = $null; summaries = @() }
            foreach ($field in @('limit', 'maxItems', 'maxDurationSeconds')) { if ($field -in @($Request.PSObject.Properties.Name)) { $browse[$field] = $Request.$field } }
            $execution.duration = $browse.maxDurationSeconds; $response.data.browse = $browse
        }
        Assert-MailLensNotCancelled $CancellationRequested
        if ($null -ne $context.Query -and [double](& $execution.Clock) - $execution.started -ge $execution.duration) { Stop-MailLens 'time_budget' 'The cooperative query time budget was reached.' }
        Write-MailLensProgress $execution 'connect'
        try {
            $execution.property = 'Connect'
            $connection = & $Adapter.Connect
            $application = $connection.Application
            $response.data.sessionMode = [string]$connection.Mode
            Assert-MailLensNotCancelled $CancellationRequested
            Write-MailLensProgress $execution 'namespace'
            $execution.property = 'GetNamespace'
            if ($Request.operation -in @('browse', 'search', 'folders')) {
                $execution.phase = 'namespace'
                $session = Invoke-MailLensBoundedCall $Adapter $application Method GetNamespace @('MAPI') $execution $CancellationRequested
            } else { $session = & $Adapter.Method $application 'GetNamespace' @('MAPI') }
            if ($null -eq $session) { throw 'Unavailable namespace.' }
        } catch { if ($_.Exception.Data['MailLensCode'] -in @('cancelled', 'time_budget')) { throw }; Stop-MailLens 'outlook_unavailable' 'Classic Outlook COM or the current profile is unavailable; allow legitimate prompts or check Classic Outlook installation.' $_.Exception }
        if ($Request.operation -in @('folders', 'search')) { $Adapter = New-MailLensBoundedAdapter $Adapter $execution $CancellationRequested }
        elseif ($Request.operation -in @('get', 'diagnose')) { $Adapter = New-MailLensDiagnosticAdapter $Adapter $execution $CancellationRequested }
        if ($Request.operation -eq 'diagnose') {
            Write-MailLensProgress $execution 'diagnose'
            $response.data.outlookVersion = $null
            try { $response.data.outlookVersion = [string](& $Adapter.Property $application 'Version') }
            catch { if ($_.Exception.Data['MailLensCode'] -eq 'cancelled') { throw }; Add-MailLensProblem $context 'version_unavailable' 'The optional Outlook version is unavailable.' }
            Assert-MailLensNotCancelled $CancellationRequested
        } elseif ($Request.operation -eq 'browse') {
            Read-MailLensBrowse $Adapter $session $Request $context $response $CancellationRequested
            Assert-MailLensNotCancelled $CancellationRequested
            if ([double](& $execution.Clock) - $execution.started -ge $execution.duration -and $response.data.browse.stopReason -ne 'time_budget') {
                $response.data.browse.stopReason = 'time_budget'; $response.data.browse.scanComplete = $false
                Add-MailLensProblem $context 'time_budget' 'The cooperative time budget was reached; collection coverage is incomplete.'
            }
        } elseif ($Request.operation -eq 'get') {
            Write-MailLensProgress $execution 'get'
            $response.data.detail = $null
            Read-MailLensDetail $Adapter $session $Request $context $response $CancellationRequested
            Assert-MailLensNotCancelled $CancellationRequested
        } else {
            $response.data.discoveryAttempted = $true
            Read-MailLensQuery $Adapter $session $Request $context $response $search $CancellationRequested
            Assert-MailLensNotCancelled $CancellationRequested
            if ([double](& $execution.Clock) - $execution.started -ge $execution.duration) { Stop-MailLens 'time_budget' 'The cooperative query time budget was reached.' }
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
        $response.errors = @([ordered]@{ code = [string]$code; message = $message; operation = $response.operation; stage = $execution.stage; property = $execution.property; hResult = $(if ($_.Exception.Data.Contains('MailLensHResult')) { $_.Exception.Data['MailLensHResult'] } else { $_.Exception.HResult }) })
        if ($code -eq 'cancelled') { $response.status = 'cancelled' } else { $response.status = 'failed' }
        if ($null -ne $context.Query -and $code -in @('cancelled', 'time_budget', 'folder_budget', 'item_budget', 'summary_limit')) {
            $context.Query.stopReason = $code; $context.Query.scanComplete = $false
            if ($null -ne $search) { $search.data.stopReason = $code; $search.data.scanComplete = $false }
            if ($code -ne 'cancelled') {
                $response.status = 'partial'; $response.errors = @()
                Add-MailLensProblem $context $code 'The requested scope was not fully evaluated because a query limit was reached.'
            }
        }

        if ($null -ne $context.Query -and $code -in @('time_budget', 'folder_budget') -and -not $context.SelectionResolved) {
            $response.status = 'failed'
            $response.errors = @([ordered]@{ code='selection_incomplete'; message='No trustworthy requested target was resolved within the query budget.'; operation=$response.operation; stage=$execution.stage; property=$execution.property; hResult=$execution.hResult })
        }
        if ($response.operation -eq 'browse' -and $response.data.Contains('browse')) {
            if ($code -in @('cancelled', 'time_budget')) { $response.data.browse.stopReason = $code; $response.data.browse.scanComplete = $false }
            if ($code -eq 'time_budget') {
                $response.status = 'partial'; $response.errors = @()
                Add-MailLensProblem $context 'time_budget' 'The cooperative time budget was reached; collection coverage is incomplete.'
            }
        }
    } finally {
        Close-MailLensReference $Adapter $session $context
        Close-MailLensReference $Adapter $application $context
        if ($null -ne $search) {
            $search.data.attemptedFolderCount = $context.FolderAttempts
            $search.data.matchedCountIsExact = ($search.data.scanComplete -and $context.WarningCount -eq 0 -and $response.status -eq 'complete')
            Complete-MailLensSearchData $search
        }
        if ($null -ne $context.Query) { $context.Query.attemptedFolderCount = $context.FolderAttempts }
        if ($null -ne $context.Query) {
            $response.scope.stores = @($context.Stores.ToArray()); $response.scope.folders = @($context.Folders.ToArray())
        }
        $response.data.stores = @($context.Stores.ToArray())
        $response.data.folders = @($context.Folders.ToArray())
        $response.data.coverage = [ordered]@{
            mountedStores = $context.StoreCount; discoveredStores = $context.Stores.Count
            readableStores = @($context.Stores.ToArray() | Where-Object { $_.readable }).Count
            discoveredFolders = $context.Folders.Count
            readableFolders = @($context.Folders.ToArray() | Where-Object { $_.readable }).Count
            complete = $(if ($response.data.discoveryAttempted) { $response.status -eq 'complete' -and $context.Problems.Count -eq 0 } else { $null })
            evaluated = $response.data.discoveryAttempted
        }
        $response.warnings = @($context.Problems.ToArray())
        if ($context.BoundedProblems) {
            $response.warningCount = $context.WarningCount
            $response.warningSamplesTruncated = ($context.WarningCount -gt $context.Problems.Count)
            $response.warningSummary = @($context.WarningGroups.Values | Sort-Object code, stage, property | ForEach-Object {
                [ordered]@{ code = $_.code; stage = $_.stage; property = $_.property; count = $_.count }
            })
        }
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
        } catch {
            $response = [ordered]@{
                schemaVersion = 2; operation = $null; status = 'failed'
                startedAt = $started.ToString('o'); endedAt = [DateTimeOffset]::UtcNow.ToString('o')
                scope = [ordered]@{ stores = @(); folders = @() }; warnings = @(); warningCount = 0; warningSummary = @(); warningSamplesTruncated = $false
                errors = @([ordered]@{ code = 'invalid_request_file'; message = 'The UTF-8 request file is unreadable or contains malformed JSON.' })
                data = $null
            }
        }
        if ($null -eq $response) { $response = Invoke-MailLensRequest $request $Adapter $Runtime $CancellationRequested $started }
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

# Managed presentation and private selection state. This is not an authentication boundary.
function Get-MailLensViewField {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return ,$property.Value }
    return $null
}

function Assert-MailLensViewShape {
    param($Object, [hashtable]$Fields)
    if ($Object -isnot [pscustomobject]) { throw 'Invalid object.' }
    foreach ($name in $Fields.Keys) {
        if ($name -cnotin @($Object.PSObject.Properties.Name)) { throw 'Missing field.' }
        $value = $Object.$name; $kind = $Fields[$name]
        if ($kind.EndsWith('?') -and $null -eq $value) { continue }
        switch ($kind.TrimEnd('?')) {
            'string' { if ($value -isnot [string]) { throw 'Invalid string.' } }
            'identity' { if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) { throw 'Invalid identity.' } }
            'int' { if (($value -isnot [int] -and $value -isnot [long]) -or $value -lt 0) { throw 'Invalid count.' } }
            'bool' { if ($value -isnot [bool]) { throw 'Invalid boolean.' } }
            'array' { if ($value -isnot [array]) { throw 'Invalid array.' } }
            'object' { if ($value -isnot [pscustomobject]) { throw 'Invalid object.' } }
            default { throw 'Invalid field specification.' }
        }
    }
}

function Select-MailLensViewFields {
    param($Object, [hashtable]$Fields)
    Assert-MailLensViewShape $Object $Fields
    $selected = [ordered]@{}
    foreach ($name in ($Fields.Keys | Sort-Object)) { $selected[$name] = $Object.$name }
    return $selected
}

function Assert-MailLensPrivatePath {
    param([string]$Path, [switch]$OutsideRepository)
    if (-not [IO.Path]::IsPathRooted($Path) -or $Path -match '^[A-Za-z]:($|[^\\/])' -or $Path -match '^\\[^\\]') { throw 'Absolute path required.' }
    $canonical = [IO.Path]::GetFullPath($Path)
    if ($canonical.Length -gt ([IO.Path]::GetPathRoot($canonical)).Length) { $canonical = $canonical.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar) }
    $cursor = $canonical
    while (-not [string]::IsNullOrEmpty($cursor)) {
        if ([IO.File]::Exists($cursor) -or [IO.Directory]::Exists($cursor)) {
            if (([IO.File]::GetAttributes($cursor) -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Reparse path unavailable.' }
        }
        if ($OutsideRepository -and ([IO.File]::Exists((Join-Path $cursor '.git')) -or [IO.Directory]::Exists((Join-Path $cursor '.git')))) { throw 'Repository path unavailable.' }
        $parent = [IO.Path]::GetDirectoryName($cursor)
        if ($parent -eq $cursor) { break }; $cursor = $parent
    }
    return $canonical
}

function Read-MailLensPrivateJson {
    param([string]$Path)
    $null = Assert-MailLensPrivatePath $Path
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { throw 'Unexpected private file BOM.' }
    return (ConvertFrom-Json -InputObject $text -ErrorAction Stop)
}

function Get-MailLensPrivateHash {
    param([string]$Path)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash([IO.File]::ReadAllBytes($Path)))).Replace('-', '').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Write-MailLensPrivateJson {
    param([string]$Path, $Value, [scriptblock]$CancellationRequested = { $false })
    $null = Assert-MailLensPrivatePath $Path
    $temporary = Join-Path ([IO.Path]::GetDirectoryName($Path)) ('.maillens-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $stream = $null
    try {
        $bytes = (New-Object Text.UTF8Encoding($false, $true)).GetBytes((ConvertTo-Json -InputObject $Value -Depth 30 -Compress))
        $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $stream.Write($bytes, 0, $bytes.Length); $stream.Dispose(); $stream = $null
        Assert-MailLensNotCancelled $CancellationRequested
        [IO.File]::Move($temporary, $Path)
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function New-MailLensOwnedRun {
    $root = Assert-MailLensPrivatePath ([IO.Path]::GetTempPath()) -OutsideRepository
    if (-not [IO.Directory]::Exists($root)) { throw 'Temporary location unavailable.' }
    $run = Join-Path $root ('mail-lens-' + [Guid]::NewGuid().ToString('N'))
    # Reserve the name before directory creation. Never adopt a preexisting run.
    $reservationPath = $run + '.maillens-lock'; $reservation = $null
    try {
        $reservation = [IO.File]::Open($reservationPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        if ([IO.Directory]::Exists($run) -or [IO.File]::Exists($run)) { throw 'Run collision.' }
        [void][IO.Directory]::CreateDirectory($run)
        $null = Assert-MailLensPrivatePath $run -OutsideRepository
        return $run
    } finally {
        if ($null -ne $reservation) { $reservation.Dispose(); [IO.File]::Delete($reservationPath) }
    }
}

function ConvertTo-MailLensCompactView {
    param($Result, [int]$ExitCode, $Request, [string]$RunDirectory)
    Assert-MailLensViewShape $Result @{ schemaVersion='int'; operation='string?'; status='string'; startedAt='string'; endedAt='string'; scope='object'; warnings='array'; errors='array'; warningCount='int'; warningSummary='array'; warningSamplesTruncated='bool'; data='object?' }
    if ($Result.schemaVersion -ne 2 -or $Result.status -cnotin @('complete','partial','failed','cancelled') -or $ExitCode -ne (Get-MailLensExitCode $Result.status)) { throw 'Envelope mismatch.' }
    $validationFailure = ($null -eq $Result.operation -and $Result.status -ceq 'failed')
    if ($validationFailure) {
        if (@($Result.errors | Where-Object { $_.code -cin @('invalid_request_file','invalid_request','unsupported_schema','unsupported_operation') }).Count -ne $Result.errors.Count) { throw 'Invalid validation diagnostic.' }
    } elseif ($Result.operation -cnotin @('browse','search','diagnose','folders','get') -or $Result.operation -cne (Get-MailLensViewField $Request 'operation')) { throw 'Operation mismatch.' }
    if ($Result.status -in @('failed','cancelled') -and $Result.errors.Count -eq 0) { throw 'Missing diagnostic.' }
    Assert-MailLensViewShape $Result.scope @{ stores='array'; folders='array' }
    $errors = @(); $groups = @()
    foreach ($errorRow in $Result.errors) {
        Assert-MailLensViewShape $errorRow @{ code='string'; message='string' }
        $safeCode = ConvertTo-MailLensDiagnosticToken $errorRow.code 'code'
        $errors += [ordered]@{ code=$safeCode; message=(Get-MailLensSafeMessage $safeCode); operation=$Result.operation
            stage=(ConvertTo-MailLensDiagnosticToken (Get-MailLensViewField $errorRow 'stage') 'stage')
            property=(ConvertTo-MailLensDiagnosticToken (Get-MailLensViewField $errorRow 'property') 'property'); hResult=(Get-MailLensSafeHResult (Get-MailLensViewField $errorRow 'hResult')) }
    }
    $warningTotal = 0
    foreach ($group in $Result.warningSummary) {
        Assert-MailLensViewShape $group @{code='string'; stage='string?'; property='string?'; count='int'}
        $groups += [ordered]@{code=(ConvertTo-MailLensDiagnosticToken $group.code 'code'); stage=(ConvertTo-MailLensDiagnosticToken $group.stage 'stage'); property=(ConvertTo-MailLensDiagnosticToken $group.property 'property'); count=$group.count}
        $warningTotal += $group.count
    }
    if ($warningTotal -ne $Result.warningCount) { throw 'Warning count mismatch.' }
    $requested = Get-MailLensViewField $Result.scope 'requested'
    if ($null -ne $requested) {Assert-MailLensViewShape $requested @{allMailboxes='bool';mailboxes='array';folders='array';includeDescendants='bool';defaultInbox='bool'}}
    if ($Result.status -cin @('complete','partial') -and $Result.operation -cin @('folders','search') -and $null -eq $requested) {throw 'Missing requested scope.'}
    $view = [ordered]@{ runDirectory=$RunDirectory; status=$Result.status; operation=$Result.operation; exitCode=$ExitCode
        requestedScope=[ordered]@{defaultInbox=($Result.operation -ceq 'browse' -and $null -eq (Get-MailLensViewField $Request 'storeId')) -or (Get-MailLensViewField $requested 'defaultInbox') -eq $true
            allMailboxes=(Get-MailLensViewField $Request 'allMailboxes') -eq $true; includeDescendants=(Get-MailLensViewField $requested 'includeDescendants') -eq $true
            selectedMailboxCount=@((Get-MailLensViewField $Request 'mailboxes') | Where-Object {$null -ne $_}).Count; selectedFolderCount=@((Get-MailLensViewField $Request 'folders') | Where-Object {$null -ne $_}).Count
            exactBrowseFolder=($Result.operation -ceq 'browse' -and $null -ne (Get-MailLensViewField $Request 'folderId'))}
        resolvedScope=[ordered]@{storeCount=$Result.scope.stores.Count; folderCount=$Result.scope.folders.Count; stores=@(); locations=@(); omittedStoreCount=[Math]::Max(0,$Result.scope.stores.Count-10); omittedLocationCount=[Math]::Max(0,$Result.scope.folders.Count-10)}
        coverage=$null; warningCount=$Result.warningCount; warningSummary=$groups; warningSamplesTruncated=$Result.warningSamplesTruncated; errors=$errors }
    foreach ($store in ($Result.scope.stores | Select-Object -First 10)) { $view.resolvedScope.stores += Select-MailLensViewFields $store @{storeName='string?';readable='bool'} }
    foreach ($folder in ($Result.scope.folders | Select-Object -First 10)) {
        $location=Select-MailLensViewFields $folder @{folderPath='string?';readable='bool'}
        $display=Get-MailLensViewField $folder 'storeName'; if ($null -ne $display -and $display -isnot [string]) {throw 'Invalid store display.'}
        $location.storeName=$display; $view.resolvedScope.locations += $location
    }
    $reliable = $Result.status -cin @('complete','partial')
    if ($reliable -or $null -ne $Result.data) {
        Assert-MailLensViewShape $Result.data @{ coverage='object'; discoveryAttempted='bool' }
        $view.coverage = Select-MailLensViewFields $Result.data.coverage @{evaluated='bool';complete='bool?';mountedStores='int?';discoveredStores='int';readableStores='int';discoveredFolders='int';readableFolders='int'}
    }
    $rows = @(); $entries = @(); $data = $null
    if ($Result.operation -cin @('browse','search','folders')) {
        $dataName = $Result.operation; if ($dataName -eq 'folders') { $dataName = 'directory' }
        $data = Get-MailLensViewField $Result.data $dataName
        if ($reliable -and $null -eq $data) { throw 'Missing operation data.' }
        if ($null -ne $data) {
            $fields = @{maxDurationSeconds='int';scanComplete='bool';stopReason='string?'}
            if ($Result.operation -eq 'browse') {
                foreach ($name in @('limit','maxItems','visitedCount','returnedCount','nonMailCount','unreadableCount')) {$fields[$name]='int'}
                $fields.ordering='string'; $fields.moreAvailable='bool?'
            } else {
                $fields.maxFolders='int'; $fields.attemptedFolderCount='int'
                if ($Result.operation -eq 'search') {
                    foreach ($name in @('summaryLimit','maxItems','visitedCount','matchedCount','returnedCount')) {$fields[$name]='int'}
                    foreach ($name in @('exhaustive','matchedCountIsExact','resultsTruncated')) {$fields[$name]='bool'}
                    $fields.ordering='string'
                }
            }
            $view.metrics = Select-MailLensViewFields $data $fields
            if ($data.stopReason -cnotin @($null,'end','limit','item_budget','time_budget','folder_budget','summary_limit','cancelled')) {throw 'Invalid stop reason.'}
            if ($Result.operation -ceq 'browse' -and $data.ordering -cnotin @('receivedTimeDescending','collectionOrder')) {throw 'Invalid browse ordering.'}
            if ($Result.operation -ceq 'search' -and $data.ordering -cne 'retainedEffectiveTimeDescending') {throw 'Invalid search ordering.'}
        }
    }
    if ($Result.operation -eq 'search' -and $null -ne $data) {
        Assert-MailLensViewShape $data @{range='object';keywords='array';keywordGroups='array';listOnly='bool';sender='string?';folderCoverage='array'}
        $view.range = Select-MailLensViewFields $data.range @{mode='string';timeZone='string';fromInclusive='string?';toExclusive='string?';localFromInclusive='string?';localToExclusive='string?'}
        foreach ($term in $data.keywords) {if ($term -isnot [string]) {throw 'Invalid term.'}}
        foreach ($group in $data.keywordGroups) { if ($group -isnot [array]) {throw 'Invalid terms.'}; foreach ($term in $group) {if ($term -isnot [string]) {throw 'Invalid term.'}} }
        $view.keywords=$data.keywords; $view.keywordGroups=$data.keywordGroups; $view.listOnly=$data.listOnly; $view.sender=$data.sender
        $view.searchCoverage=[ordered]@{attemptedRows=$data.folderCoverage.Count;incompleteFolders=0;unreadableCount=0;timestampClassificationGaps=0}
        foreach ($folder in $data.folderCoverage) {
            Assert-MailLensViewShape $folder @{complete='bool';unreadableCount='int';timestampClassificationComplete='bool'}
            if (-not $folder.complete) {$view.searchCoverage.incompleteFolders++}; if (-not $folder.timestampClassificationComplete) {$view.searchCoverage.timestampClassificationGaps++}
            $view.searchCoverage.unreadableCount += $folder.unreadableCount
        }
    }
    if ($Result.operation -cin @('browse','search')) {
        if ($null -ne $data) { Assert-MailLensViewShape $data @{summaries='array'}; $rows=@($data.summaries) }
        $view.selectedCount=0; $view.omittedSummaryCount=$rows.Count; $view.summaries=@()
        if ($reliable) {
            if ($data.returnedCount -ne $rows.Count) {throw 'Summary count mismatch.'}
            for ($position=1; $position -le [Math]::Min(10,$rows.Count); $position++) {
                $row=$rows[$position-1]
                Assert-MailLensViewShape $row @{storeId='identity';entryId='identity';folderId='identity';subject='string';timestamp='string?';timestampField='string?';storeName='string?';folderPath='string?'}
                $label='M'+$position
                $summary=Select-MailLensViewFields $row @{subject='string';timestamp='string?';timestampField='string?';storeName='string?';folderPath='string?'}
                foreach ($flag in @('timestampClassificationComplete','consistencyAvailable','contentChangedDuringRead')) {
                    if ($flag -cin @($row.PSObject.Properties.Name)) {Assert-MailLensViewShape $row @{$flag='bool'}; $summary[$flag]=$row.$flag}
                }
                $summary.label=$label; $view.summaries += $summary
                $entries += [ordered]@{label=$label;summaryPosition=$position;storeId=$row.storeId;entryId=$row.entryId}
            }
            $view.selectedCount=$entries.Count; $view.omittedSummaryCount=$rows.Count-$entries.Count
        }
    }
    if ($Result.operation -ceq 'get' -and $reliable) {
        $detail=Get-MailLensViewField $Result.data 'detail'
        $view.detail=Select-MailLensViewFields $detail @{body='string';bodySource='string';bodyOffsetUnit='string';bodyTotalLength='int';bodyOffset='int';bodyLength='int';bodyReturnedLength='int';bodyTruncated='bool';lastModificationTime='string?';lastModificationTimeAfter='string?';consistencyAvailable='bool';contentChangedDuringRead='bool';sentOn='string?';receivedTime='string?';creationTime='string?';recipientsComplete='bool';attachmentNamesComplete='bool'}
        Assert-MailLensViewShape $detail @{source='object';sender='object';recipients='array';attachmentNames='array'}
        $view.detail.source=Select-MailLensViewFields $detail.source @{subject='string?';storeName='string?';folderPath='string?';timestamp='string?';timestampField='string?'}
        Assert-MailLensViewShape $detail.source @{storeId='identity';entryId='identity';folderId='identity'}
        if ($detail.source.storeId -cne $Request.storeId -or $detail.source.entryId -cne $Request.entryId -or $detail.bodySource -cne 'Body' -or $detail.bodyOffsetUnit -cne 'unicodeScalarValues') {throw 'Detail identity/source mismatch.'}
        $expectedOffset=0; $expectedLength=2000
        if ($null -ne (Get-MailLensViewField $Request 'bodyOffset')) {$expectedOffset=$Request.bodyOffset}; if ($null -ne (Get-MailLensViewField $Request 'bodyLength')) {$expectedLength=$Request.bodyLength}
        if ($detail.bodyOffset -ne $expectedOffset -or $detail.bodyLength -ne $expectedLength) {throw 'Detail chunk mismatch.'}
        $chunk=Read-MailLensBodyChunk $detail.body 0 100000 { $false }
        if ($chunk.bodyTotalLength -ne $detail.bodyReturnedLength -or $detail.bodyReturnedLength -ne [Math]::Min($detail.bodyLength,[Math]::Max(0,$detail.bodyTotalLength-$detail.bodyOffset)) -or $detail.bodyTruncated -ne ([long]$detail.bodyOffset+$detail.bodyReturnedLength -lt $detail.bodyTotalLength)) {throw 'Invalid chunk.'}
        $view.detail.sender=Select-MailLensViewFields $detail.sender @{displayName='string?';rawAddress='string?';addressType='string?'}
        $view.detail.recipients=@(); foreach ($recipient in $detail.recipients) {$view.detail.recipients += Select-MailLensViewFields $recipient @{displayName='string?';rawAddress='string?';addressType='string?';recipientType='int?'}}
        foreach ($name in $detail.attachmentNames) {if ($name -isnot [string]) {throw 'Invalid attachment name.'}}
        $view.detail.attachmentNames=$detail.attachmentNames
    }
    return @{view=$view;entries=$entries}
}

function Read-MailLensSelection {
    param([string]$RunDirectory, [string]$Label)
    if ($Label -cnotmatch '^M([1-9]|10)$') {throw 'Invalid label.'}
    $run=Assert-MailLensPrivatePath $RunDirectory -OutsideRepository
    if (-not [IO.Directory]::Exists($run)) {throw 'Missing run.'}
    $index=Read-MailLensPrivateJson (Join-Path $run 'index.json')
    Assert-MailLensViewShape $index @{formatVersion='int';runDirectory='string';resultFile='string';resultSha256='string';schemaVersion='int';operation='string';status='string';entries='array'}
    if ($index.formatVersion -ne 1 -or $index.schemaVersion -ne 2 -or $index.resultFile -cne 'result.json' -or $index.runDirectory -cne $run -or $index.operation -cnotin @('browse','search') -or $index.status -cnotin @('complete','partial') -or $index.resultSha256 -cnotmatch '^[0-9a-f]{64}$') {throw 'Unbound index.'}
    $path=Join-Path $run 'result.json'; $null=Assert-MailLensPrivatePath $path
    if ((Get-MailLensPrivateHash $path) -cne $index.resultSha256) {throw 'Stale result.'}
    $raw=Read-MailLensPrivateJson $path
    $request=[pscustomobject]@{operation=$index.operation}
    $projection=ConvertTo-MailLensCompactView $raw (Get-MailLensExitCode $index.status) $request $run
    if ($raw.status -cne $index.status -or $index.entries.Count -ne $projection.entries.Count) {throw 'Index mismatch.'}
    $selected=$null
    for ($i=0; $i -lt $index.entries.Count; $i++) {
        $entry=$index.entries[$i]; $expected=$projection.entries[$i]
        Assert-MailLensViewShape $entry @{label='string';summaryPosition='int';storeId='identity';entryId='identity'}
        foreach ($field in @('label','summaryPosition','storeId','entryId')) {if ($entry.$field -cne $expected[$field]) {throw 'Index association mismatch.'}}
        if ($entry.label -ceq $Label) {$selected=$entry}
    }
    if ($null -eq $selected) {throw 'Unknown label.'}
    return $selected
}

function Write-MailLensCompactStdout {
    param([string]$Json)
    $bytes=(New-Object Text.UTF8Encoding($false,$true)).GetBytes($Json+[Environment]::NewLine)
    $stream=[Console]::OpenStandardOutput()
    $stream.Write($bytes,0,$bytes.Length); $stream.Flush()
}

function Invoke-MailLensManagedRequest {
    [CmdletBinding()]
    param([string]$RequestPath, [string]$RunDirectory, [string]$Label, $BodyOffset=0, $BodyLength=2000,
        $Adapter=(New-MailLensComAdapter), $Runtime=(Get-MailLensRuntime), [scriptblock]$CancellationRequested={ $false },
        [scriptblock]$WriteView={param($Json) Write-MailLensCompactStdout $Json})
    $run=$null
    try {
        $labelMode=$PSBoundParameters.ContainsKey('RunDirectory') -or $PSBoundParameters.ContainsKey('Label')
        if ($labelMode) {
            if ($PSBoundParameters.ContainsKey('RequestPath') -or -not $RunDirectory -or -not $Label -or $BodyOffset -isnot [int] -or $BodyOffset -lt 0 -or $BodyLength -isnot [int] -or $BodyLength -lt 1 -or $BodyLength -gt 100000) {throw 'Invalid managed mode.'}
            $entry=Read-MailLensSelection $RunDirectory $Label
            $request=[pscustomobject]@{schemaVersion=2;operation='get';storeId=$entry.storeId;entryId=$entry.entryId;bodyOffset=$BodyOffset;bodyLength=$BodyLength}
            $originalRun=Assert-MailLensPrivatePath $RunDirectory -OutsideRepository
        } else {
            if (-not $RequestPath -or $PSBoundParameters.ContainsKey('BodyOffset') -or $PSBoundParameters.ContainsKey('BodyLength')) {throw 'Invalid managed mode.'}
            $source=Assert-MailLensPrivatePath $RequestPath -OutsideRepository
        }
        Assert-MailLensNotCancelled $CancellationRequested
        $run=New-MailLensOwnedRun
        $requestFile=Join-Path $run 'request.json'; $resultFile=Join-Path $run 'result.json'
        if ($labelMode) {Write-MailLensPrivateJson $requestFile $request $CancellationRequested}
        else {
            # Byte-copy preserves optional request BOM and raw parse/validation failure semantics.
            $inputBytes=[IO.File]::ReadAllBytes($source)
            $stream=[IO.File]::Open($requestFile,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
            try {$stream.Write($inputBytes,0,$inputBytes.Length)} finally {$stream.Dispose()}
            $request=$null
            try {
                $text=(New-Object Text.UTF8Encoding($false,$true)).GetString($inputBytes)
                if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) {$text=$text.Substring(1)}
                $request=ConvertFrom-Json -InputObject $text -ErrorAction Stop
            } catch { $request=$null }
        }
        $code=Invoke-MailLensFileRequest $requestFile $resultFile $Adapter $Runtime $CancellationRequested
        $result=Read-MailLensPrivateJson $resultFile
        $projection=ConvertTo-MailLensCompactView $result $code $request $run
        if ($labelMode) {$projection.view.label=$Label; $projection.view.selectionRunDirectory=$originalRun}
        # Construct/encode the entire view before committing a usable index or emitting success.
        $json=ConvertTo-Json -InputObject $projection.view -Depth 30 -Compress
        $null=(New-Object Text.UTF8Encoding($false,$true)).GetBytes($json)
        if ($result.operation -cin @('browse','search') -and $result.status -cin @('complete','partial')) {
            $index=[ordered]@{formatVersion=1;runDirectory=$run;resultFile='result.json';resultSha256=(Get-MailLensPrivateHash $resultFile);schemaVersion=2;operation=$result.operation;status=$result.status;entries=@($projection.entries)}
            Write-MailLensPrivateJson (Join-Path $run 'index.json') $index $CancellationRequested
        }
        & $WriteView $json
        return $code
    } catch {
        $cancelled=$_.Exception.Data['MailLensCode'] -eq 'cancelled'
        $code=1; $status='failed'; $errorCode='caller_unavailable'; $message='MailLens caller state or output is unavailable; coverage is unknown.'
        if ($cancelled) {$code=130; $status='cancelled'; $errorCode='cancelled'; $message='MailLens caller was cancelled before completion.'}
        $failure=[ordered]@{runDirectory=$run;status=$status;operation=$null;exitCode=$code;errors=@(@{code=$errorCode;message=$message})}
        try {& $WriteView (ConvertTo-Json -InputObject $failure -Depth 5 -Compress)} catch { [Console]::Error.WriteLine('MailLens compact output could not be delivered.') }
        return $code
    }
}

function Get-MailLensSafeMessage {
 param([string]$Code)
 $messages=@{
  'invalid_request'='The request contains an invalid field or value.'
  'unsupported_schema'='schemaVersion must be the integer 2; update the complete package and request together.'
  'unsupported_operation'='operation must be diagnose, browse, folders, search or get.'
  'time_budget'='The cooperative time budget was reached; collection coverage is incomplete.'
  'folder_unavailable'='The exact requested folder is unavailable.'
  'default_store_unavailable'='The existing default store is unavailable.'
  'default_inbox_unavailable'='The existing default Inbox is unavailable.'
  'folder_identity_unavailable'='An attempted folder identity is unavailable.'
  'display_metadata_unavailable'='The selected folder store metadata is unavailable.'
  'items_unavailable'='The selected folder Items collection is unavailable.'
  'sort_unavailable'='Sorting is unavailable; this is a bounded collection-order sample, not the latest messages.'
  'timestamp_unavailable'='The optional ReceivedTime display timestamp is unavailable.'
  'browse_item_unavailable'='An attempted item has unavailable identity or subject metadata.'
  'item_budget'='The global item budget was reached.'
  'release_failed'='An acquired Outlook reference could not be released.'
  'cancelled'='The operation was cancelled.'
  'folder_budget'='The global folder budget was reached; requested scope is incomplete.'
  'store_identity_unavailable'='The selected store identity is unavailable.'
  'folder_denied'='The selected child-folder collection is unavailable.'
  'store_metadata_unavailable'='Mounted-store metadata is incomplete; unresolved selectors cannot be widened.'
  'selection_incomplete'='No trustworthy requested target was resolved within the query budget.'
  'unknown_selection'='The requested path was not found within the selected stores.'
  'ambiguous_selection'='The requested path is ambiguous; use an exact storeId/folderId pair.'
  'folder_identity_changed'='The exact requested folder identity changed.'
  'store_denied'='An attempted selected store or root is unavailable.'
  'selection_unavailable'='No trustworthy folder target could be resolved.'
  'sent_folder_denied'='Sent Items classification is unavailable; fallback timestamps cannot certify exact time coverage.'
  'sent_ancestry_unavailable'='Selected-folder ancestry is unavailable; fallback timestamps are qualified.'
  'date_dst_gap'='An Outlook local timestamp is ambiguous or nonexistent in the Windows timezone.'
  'date_unavailable'='No valid effective timestamp is available for this mail item.'
  'consistency_unavailable'='LastModificationTime is unavailable; readable data is not a certified consistent snapshot.'
  'date_filter_unavailable'='Date Restrict is unavailable; this folder uses the original collection under the global budgets.'
  'sender_denied'='Sender metadata is partly unavailable; narrowing coverage is incomplete.'
  'summary_limit'='The bounded summary limit was reached; remaining scope is unknown.'
  'folder_capability_denied'='Repeated collection access denial closed this folder path; remaining items are unknown.'
  'search_folder_denied'='The selected item collection is unavailable or changed; matching coverage is incomplete.'
  'detail_property_unavailable'='The current store display metadata is unavailable.'
  'recipient_unavailable'='Recipient address type is unavailable; unresolved names are not resolved.'
  'detail_collection_unavailable'='A requested detail collection is unavailable or changed.'
  'body_unavailable'='The requested plain Body is unavailable; it is not an empty body.'
  'item_unavailable'='The exact storeId/entryId item is stale, deleted, moved or inaccessible.'
  'unsupported_item_type'='The requested item is not a MailItem.'
  'item_identity_unavailable'='The requested item type or exact source identity could not be verified.'
  'item_identity_changed'='The item source changed during retrieval; reread the current identity.'
  'item_content_changed'='The item changed during retrieval; discard combined chunks and reread.'
  'unsupported_platform'='MailLens requires native Windows.'
  'unsupported_powershell'='MailLens requires Windows PowerShell 5.1.'
  'unsupported_apartment'='Launch powershell.exe with -NoProfile -STA.'
  'outlook_unavailable'='Classic Outlook COM or the current profile is unavailable; allow legitimate prompts or check Classic Outlook installation.'
  'version_unavailable'='The optional Outlook version is unavailable.'
  'invalid_request_file'='The UTF-8 request file is unreadable or contains malformed JSON.'
  'discovery_failed'='Outlook discovery failed; coverage is incomplete.'
 }
 if ($messages.ContainsKey($Code)) {return $messages[$Code]}
 return 'MailLens query could not complete; inspect the private diagnostics.'
}
function Get-MailLensSafeHResult {
 param($Value)
 if ($null -eq $Value) {return $null}
 if (($Value -isnot [int] -and $Value -isnot [long]) -or $Value -lt [int]::MinValue -or $Value -gt [int]::MaxValue) {throw 'Invalid HResult.'}
 return $Value
}
function ConvertTo-MailLensDiagnosticToken {
 param($Value,[string]$Kind)
 if ($null -eq $Value -or $Value -ceq '') {return $Value}
 $allowed=@{code=@('cancelled','ambiguous_selection','body_unavailable','browse_item_unavailable','consistency_unavailable','date_dst_gap','date_filter_unavailable','date_unavailable','default_inbox_unavailable','default_store_unavailable','detail_collection_unavailable','detail_property_unavailable','discovery_failed','display_metadata_unavailable','folder_budget','folder_capability_denied','folder_denied','folder_identity_changed','folder_identity_unavailable','folder_unavailable','invalid_request','invalid_request_file','item_budget','item_content_changed','item_identity_changed','item_identity_unavailable','item_unavailable','items_unavailable','outlook_unavailable','recipient_unavailable','release_failed','search_folder_denied','selection_incomplete','selection_unavailable','sender_denied','sent_ancestry_unavailable','sent_folder_denied','sort_unavailable','store_denied','store_identity_unavailable','store_metadata_unavailable','summary_limit','time_budget','timestamp_unavailable','unknown_selection','unsupported_apartment','unsupported_item_type','unsupported_operation','unsupported_platform','unsupported_powershell','unsupported_schema','version_unavailable');stage=@('acquire-folder','acquire-items','browse-items','classify-timestamp','connect','diagnose','discover-folders','get','namespace','resolve-folder','resolve-scope','restrict-date','search-items','sort-items');property=@('Address','AddressEntry','Attachments','Body','BoundedProblems','Class','Connect','Count','CreationTime','DefaultItemType','DefaultStore','Desktop','DisplayName','EntryID','FileName','FolderPath','Folders','GetDefaultFolder','GetFolderFromID','GetItemFromID','GetNamespace','GetRootFolder','Item','Items','LastModificationTime','MAPI','MailLensCode','MailLensHResult','Name','PSEdition','Parent','Property','ReceivedTime','Recipients','Resolved','Restrict','STA','SenderEmailAddress','SenderEmailType','SenderName','Sent','SentOn','Sort','Store','StoreID','Stores','Subject','Type','Version')}
 if ($Value -isnot [string]) {throw 'Invalid diagnostic.'}
 if ($Value -cnotin $allowed[$Kind]) {return 'unavailable'}
 return $Value
}

Export-ModuleMember -Function New-MailLensComAdapter, Get-MailLensRuntime, Invoke-MailLensRequest, Invoke-MailLensFileRequest, Get-MailLensExitCode, Invoke-MailLensManagedRequest
