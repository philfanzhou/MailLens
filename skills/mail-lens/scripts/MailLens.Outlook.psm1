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
    # Search/detail handlers deliberately remain outside this first session slice.
    if ($Request.operation -in @('search', 'get')) {
        Stop-MailLens 'unsupported_operation' 'This operation is not implemented in this version.'
    }
    $allowed = @('schemaVersion', 'operation')
    if ($Request.operation -eq 'folders') { $allowed += @('mailboxes', 'folders', 'includeDescendants') }
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
    $application = $null; $session = $null
    try {
        Assert-MailLensRequest $Request
        $response.operation = $Request.operation
        if (-not $Runtime.isWindows) { Stop-MailLens 'unsupported_platform' 'MailLens requires native Windows.' }
        if ($Runtime.edition -cne 'Desktop' -or $Runtime.powerShellVersion -notmatch '^5\.1\.') { Stop-MailLens 'unsupported_powershell' 'MailLens requires Windows PowerShell 5.1.' }
        if ($Runtime.apartmentState -cne 'STA') { Stop-MailLens 'unsupported_apartment' 'Launch powershell.exe with -NoProfile -STA.' }
        Assert-MailLensNotCancelled $CancellationRequested
        try {
            $connection = & $Adapter.Connect
            $application = $connection.Application
            $response.data.sessionMode = [string]$connection.Mode
            $session = & $Adapter.Method $application 'GetNamespace' @('MAPI')
        } catch { Stop-MailLens 'outlook_unavailable' 'Classic Outlook COM or the current profile is unavailable; allow legitimate prompts or check Classic Outlook installation.' }
        Read-MailLensCatalog $Adapter $session $context $CancellationRequested
        $response.scope = Resolve-MailLensScope $Request $context
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
