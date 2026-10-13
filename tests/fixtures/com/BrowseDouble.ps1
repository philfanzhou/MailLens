# Synthetic browse data; only the specified read-only acquisition methods are supported.
function New-BrowseDouble {
    param([int]$Count = 1000, [hashtable]$Failures = @{})
    $double = New-SessionDouble -Failures $Failures
    $state = $double.state
    $store = $state.application.methods.GetNamespace.properties.Stores.methods.Item[0]
    $state.application.methods.GetNamespace.properties.DefaultStore = $store
    $values = @()
    for ($i = 1; $i -le $Count; $i++) {
        $values += [MailLens.Tests.Reference]::new(('browse-' + $i), 'mail', @{
            Class = 43; EntryID = ('browse-' + $i); Subject = ('合成📮 ' + $i)
            ReceivedTime = [DateTime]::SpecifyKind(([DateTime]'2026-10-01').AddMinutes($i), [DateTimeKind]::Utc)
        }, @{})
    }
    $items = [MailLens.Tests.Reference]::new('inbox-items', 'items', @{ Count = $Count }, @{ Item = $values; Sort = $null })
    $folder = [MailLens.Tests.Reference]::new('inbox', 'folder', @{ StoreID = 's-a'; EntryID = 'inbox'; FolderPath = '\\Synthetic\Inbox'; Store = $store; Items = $items }, @{})
    foreach ($item in $values) { $item.properties.Parent = $folder }
    $store.methods.GetDefaultFolder = $folder
    $state.folder = $folder; $state.items = $items; $state.messages = $values
    $state.actions = @{}
    $baseMethod = $double.adapter.Method
    $double.adapter.Method = {
        param($Reference, $Name, $Arguments)
        $key = $Reference.refId + '.' + $Name
        if ($Name -in @('Sort', 'GetFolderFromID')) {
            [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic private denied method.' }
            if ($Name -eq 'Sort') {
                if ($Arguments[0] -cne '[ReceivedTime]' -or $Arguments[1] -ne $true) { throw 'Unexpected sort.' }
                $Reference.methods.Item = @($Reference.methods.Item | Sort-Object { $_.properties.ReceivedTime } -Descending)
                return
            }
            if ($Arguments[0] -cne 'inbox' -or $Arguments[1] -cne 's-a') { throw 'Unknown exact pair.' }
            [void]$state.acquired.Add($state.folder.refId)
            return ,$state.folder
        }
        if ($Name -eq 'GetDefaultFolder' -and $Arguments[0] -ne 6) { throw 'Only existing Inbox lookup is allowed.' }
        if ($state.actions.ContainsKey($key)) { & $state.actions[$key] $Reference }
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
