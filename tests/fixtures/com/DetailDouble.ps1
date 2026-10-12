# Synthetic content only. Reuse the fail-closed adapter and acquisition accounting.
function New-DetailDouble {
    param([string]$Body = "合同📮`r`nline", [string[]]$AttachmentNames = @(), [hashtable]$Failures = @{})
    $mail = New-SyntheticMail -Body $Body
    $other = New-SyntheticMail -Body 'Other store, same subject'
    $double = New-SearchDouble @{ 'f-root' = @($mail); 'f-archive' = @($other) } $Failures
    $state = $double.state
    $state.mail = $mail; $state.other = $other
    foreach ($folder in $state.folderMap.Values) {
        $folder.properties.Store = $state.stores[$(if ($folder.properties.StoreID -eq 's-a') { 0 } else { 1 })]
    }
    $attachmentRows = @(); $index = 0
    foreach ($name in $AttachmentNames) {
        $index++
        $attachmentRows += [MailLens.Tests.Reference]::new(('attachment-' + $index), 'attachment', @{ FileName = $name }, @{})
    }
    $mail.properties.Attachments = [MailLens.Tests.Reference]::new('attachments', 'attachments', @{ Count = $attachmentRows.Count }, @{ Item = $attachmentRows })
    $address = [MailLens.Tests.Reference]::new('recipient-address', 'address', @{ Type = 'EX' }, @{})
    $recipient = [MailLens.Tests.Reference]::new('recipient-1', 'recipient', @{ Name = 'Synthetic Recipient'; Address = '/O=Synthetic/CN=Recipient'; Type = 1; Resolved = $true; AddressEntry = $address }, @{})
    $mail.properties.Recipients = [MailLens.Tests.Reference]::new('recipients', 'recipients', @{ Count = 1 }, @{ Item = @($recipient) })
    $other.properties.Recipients = [MailLens.Tests.Reference]::new('other-recipients', 'recipients', @{ Count = 0 }, @{ Item = @() })
    $other.properties.Attachments = [MailLens.Tests.Reference]::new('other-attachments', 'attachments', @{ Count = 0 }, @{ Item = @() })
    $state.lookup = @{ 's-a|m-1' = $mail; 's-b|m-1' = $other }
    $baseMethod = $double.adapter.Method
    $double.adapter.Method = {
        param($Reference, $Name, $Arguments)
        if ($Name -eq 'GetItemFromID') {
            [void]$state.calls.Add('session.GetItemFromID.' + $Arguments[1] + '.' + $Arguments[0])
            if ($state.failures.ContainsKey('session.GetItemFromID')) { throw 'Synthetic lookup denied: never publish this body.' }
            $value = $state.lookup[$Arguments[1] + '|' + $Arguments[0]]
            if ($null -eq $value) { throw 'Synthetic stale identity.' }
            [void]$state.acquired.Add($value.refId)
            return ,$value
        }
        & $baseMethod $Reference $Name $Arguments
    }.GetNewClosure()
    return $double
}
