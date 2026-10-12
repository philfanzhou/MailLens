# Synthetic message/store identities and content only. The base double rejects all writes.
function New-SearchDouble {
    param([hashtable]$Messages = @{}, [hashtable]$Failures = @{})
    $double = New-SessionDouble -Failures $Failures
    $state = $double.state
    $state.folderMap = @{}; $state.actions = @{}
    $pending = New-Object 'System.Collections.Generic.List[object]'
    [void]$pending.Add($state.root); [void]$pending.Add($state.archiveRoot)
    for ($i = 0; $i -lt $pending.Count; $i++) {
        $folder = $pending[$i]
        if ($state.folderMap.ContainsKey($folder.refId)) { continue }
        $state.folderMap[$folder.refId] = $folder
        $folder.properties.StoreID = $(if ($folder.refId -eq 'f-archive') { 's-b' } else { 's-a' })
        $values = @(); if ($Messages.ContainsKey($folder.refId)) { $values = @($Messages[$folder.refId]) }
        foreach ($item in $values) { $item.properties.Parent = $folder }
        $folder.properties.Items = [MailLens.Tests.Reference]::new(($folder.refId + '-items'), 'items', @{ Count = $values.Count }, @{ Item = $values })
        foreach ($child in $folder.properties.Folders.methods.Item) { [void]$pending.Add($child) }
    }
    $stores = $state.application.methods.GetNamespace.properties.Stores.methods.Item
    foreach ($store in $stores) { $store.methods.GetDefaultFolder = $null }
    $state.stores = $stores
    $baseMethod = $double.adapter.Method
    $double.adapter.Method = {
        param($Reference, $Name, $Arguments)
        if ($Name -in @('GetFolderFromID', 'GetDefaultFolder')) {
            $key = $Reference.refId + '.' + $Name
            if ($Name -eq 'GetFolderFromID') { $key += '.' + $Arguments[0] }
            [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic method denied.' }
            if ($Name -eq 'GetFolderFromID') {
                $value = $state.folderMap[$Arguments[0]]
                if ($null -eq $value -or $value.properties.StoreID -cne $Arguments[1]) { throw 'Unknown folder identity.' }
            } else {
                if ($Arguments[0] -ne 5) { throw 'Only Sent Items lookup is allowed.' }
                $value = $Reference.methods.GetDefaultFolder
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
