# Synthetic message/store identities and content only. The base double rejects all writes.
function New-SearchDouble {
    param([hashtable]$Messages = @{}, [hashtable]$Failures = @{}, [switch]$SeparateInbox)
    $double = New-SessionDouble -Failures $Failures
    $state = $double.state
    $state.actions = @{}; $state.filters = New-Object 'System.Collections.Generic.List[string]'
    if ($SeparateInbox) {
        $store = $state.root.properties.Store
        $children = [MailLens.Tests.Reference]::new('inbox-children','collection',@{Count=0},@{Item=@()})
        $inbox = [MailLens.Tests.Reference]::new('f-inbox','folder',@{StoreID='s-a';EntryID='f-inbox';FolderPath='\\Synthetic\Inbox';DefaultItemType=0;Folders=$children;Store=$store;Parent=$state.root},@{})
        $state.folderMap['f-inbox']=$inbox; $store.methods.GetDefaultFolder=$inbox
        $state.root.properties.Folders.methods.Item = @($inbox) + @($state.root.properties.Folders.methods.Item)
        $state.root.properties.Folders.properties.Count++
    }
    $pending = New-Object 'System.Collections.Generic.List[object]'
    [void]$pending.Add($state.root); [void]$pending.Add($state.archiveRoot)
    $populated = @{}
    for ($i = 0; $i -lt $pending.Count; $i++) {
        $folder = $pending[$i]
        if ($populated.ContainsKey($folder.refId)) { continue }; $populated[$folder.refId] = $true
        $state.folderMap[$folder.refId] = $folder
        $folder.properties.StoreID = $(if ($folder.refId -eq 'f-archive') { 's-b' } else { 's-a' })
        $values = @(); if ($Messages.ContainsKey($folder.refId)) { $values = @($Messages[$folder.refId]) }
        foreach ($item in $values) { $item.properties.Parent = $folder }
        $folder.properties.Items = [MailLens.Tests.Reference]::new(($folder.refId + '-items'), 'items', @{ Count = $values.Count }, @{ Item = $values })
        foreach ($child in $folder.properties.Folders.methods.Item) { [void]$pending.Add($child) }
    }
    $stores = $state.application.methods.GetNamespace.properties.Stores.methods.Item
    foreach ($store in $stores) { $store.methods.GetSentFolder = $null }
    $state.stores = $stores; $state.timeZone = 'UTC'
    $baseMethod = $double.adapter.Method
    $double.adapter.Method = {
        param($Reference, $Name, $Arguments)
        if ($Name -eq 'Restrict') {
            $key = $Reference.refId + '.Restrict'; [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic unsupported date filter.' }
            $filter = [string]$Arguments[0]; [void]$state.filters.Add($filter)
            # Parse every clause of the emitted predicate; no candidates are added
            # just because the exact evaluator would consider their dates invalid.
            $clause = '\("(?<property>[^"\r\n]+)" >= ''(?<from>(?:[^'']|'''')*)'' AND "\k<property>" < ''(?<to>(?:[^'']|'''')*)''\) OR ' +
                '\("\k<property>" IS NULL OR "\k<property>" <= ''(?<invalidLower>(?:[^'']|'''')*)'' OR "\k<property>" >= ''(?<invalidUpper>(?:[^'']|'''')*)''\)'
            $clauses = [regex]::Matches($filter, $clause)
            if ($clauses.Count -ne 3 -or ('@SQL=' + (($clauses | ForEach-Object { $_.Value }) -join ' OR ')) -cne $filter) { throw 'Unexpected date union.' }
            $fields = @{'urn:schemas:httpmail:datereceived'='ReceivedTime';'urn:schemas:httpmail:date'='SentOn';'http://schemas.microsoft.com/mapi/proptag/0x30070040'='CreationTime'}
            $culture = [Globalization.CultureInfo]::CurrentCulture
            $zone = [TimeZoneInfo]::FindSystemTimeZoneById($state.timeZone)
            $values = @($Reference.methods.Item | Where-Object {
                $mail = $_; $included = $false
                foreach ($part in $clauses) {
                    $field = $fields[$part.Groups['property'].Value]
                    if ($null -eq $field) { throw 'Unexpected date property.' }
                    $date = $mail.properties[$field]
                    # Outlook exposes an unset (null) date as its 4501 sentinel.
                    if ($null -eq $date -or $date -eq [DateTime]'4501-01-01') { $included = $true; continue }
                    if ($date -isnot [DateTime]) { throw 'Synthetic provider date must be DateTime or null.' }
                    if ($date.Kind -ne [DateTimeKind]::Utc) {
                        # Synthetic provider storage uses a deterministic offset;
                        # exact evaluation still rejects ambiguous/nonexistent COM dates.
                        $local = [DateTime]::SpecifyKind($date, [DateTimeKind]::Unspecified)
                        if ($date -gt [DateTime]'1601-01-01' -and $date.Year -lt 4501) {
                            $date = [DateTime]::SpecifyKind($local.Subtract($zone.GetUtcOffset($local)), [DateTimeKind]::Utc)
                        }
                    }
                    $bounds = @{}
                    foreach ($name in @('from','to','invalidLower','invalidUpper')) {
                        $bounds[$name] = [DateTime]::SpecifyKind([DateTime]::Parse($part.Groups[$name].Value.Replace("'" + "'", "'"), $culture), [DateTimeKind]::Utc)
                    }
                    if (($date -ge $bounds.from -and $date -lt $bounds.to) -or $date -le $bounds.invalidLower -or $date -ge $bounds.invalidUpper) { $included = $true }
                }
                $included
            })
            $value = [MailLens.Tests.Reference]::new(($Reference.refId + '-restricted'), 'items', @{ Count = $values.Count }, @{ Item = $values })
            [void]$state.acquired.Add($value.refId); return ,$value
        }
        if ($Name -in @('GetFolderFromID', 'GetDefaultFolder')) {
            $key = $Reference.refId + '.' + $Name
            if ($Name -eq 'GetFolderFromID') { $key += '.' + $Arguments[0] }
            [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic method denied.' }
            if ($Name -eq 'GetFolderFromID') {
                $value = $state.folderMap[$Arguments[0]]
                if ($null -eq $value -or $value.properties.StoreID -cne $Arguments[1]) { throw 'Unknown folder identity.' }
            } else {
                if ($Arguments[0] -eq 5) { $value = $Reference.methods.GetSentFolder }
                elseif ($Arguments[0] -eq 6) { $value = $Reference.methods.GetDefaultFolder }
                else { throw 'Unexpected default folder.' }
            }
            if ($null -ne $value) { [void]$state.acquired.Add($value.refId) }
            return ,$value
        }
        & $baseMethod $Reference $Name $Arguments
    }.GetNewClosure()
    $baseProperty = $double.adapter.Property
    $double.adapter.Property = {
        param($Reference, $Name)
        $key = $Reference.refId + '.' + $Name
        if ($state.actions.ContainsKey($key)) { & $state.actions[$key] $Reference }
        & $baseProperty $Reference $Name
    }.GetNewClosure()
    return $double
}

function New-SyntheticMail {
    param([string]$Id = 'm-1', [string]$Subject = 'Project A', [string]$Body = '合同 agreement',
        [DateTime]$Received = [DateTime]::SpecifyKind([DateTime]'2026-10-01T12:00:00', [DateTimeKind]::Utc),
        [bool]$Sent = $true)
    [MailLens.Tests.Reference]::new($Id, 'mail', @{
        Class = 43; EntryID = $Id; Subject = $Subject; Body = $Body; Sent = $Sent
        ReceivedTime = $Received; SentOn = $Received; CreationTime = $Received
        LastModificationTime = $Received; SenderName = 'Synthetic Sender'; SenderEmailAddress = '/O=Synthetic/OU=Test/CN=Sender'; SenderEmailType = 'EX'
    }, @{})
}
