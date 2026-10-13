# Synthetic identities only. Unexpected properties/methods fail closed.
if (-not ('MailLens.Tests.Reference' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Collections;
namespace MailLens.Tests {
    public sealed class Reference {
        public string refId { get; private set; }
        public string kind { get; private set; }
        public Hashtable properties { get; private set; }
        public Hashtable methods { get; private set; }
        public Reference(string id, string type, Hashtable values, Hashtable calls) {
            refId = id; kind = type; properties = values; methods = calls;
        }
    }
}
'@
}
function New-SessionDouble {
    param([string]$Mode = 'existing', [hashtable]$Failures = @{}, [switch]$Empty, [switch]$DuplicateNames)
    $state = @{
        acquired = New-Object 'System.Collections.Generic.List[string]'
        released = New-Object 'System.Collections.Generic.List[string]'
        calls = New-Object 'System.Collections.Generic.List[string]'
        failures = $Failures; mode = $Mode; connections = 0
    }
    function Node($id, $kind, $properties, $methods) {
        [MailLens.Tests.Reference]::new($id, $kind, $properties, $methods)
    }
    function Folder($id, $path, $type, $children) {
        $collection = Node ($id + '-children') 'collection' @{ Count = $children.Count } @{ Item = $children }
        Node $id 'folder' @{ EntryID = $id; FolderPath = $path; DefaultItemType = $type; Folders = $collection } @{}
    }
    $nested = Folder 'f-nested' '\\工作📮\Projects\合同' 0 @()
    $projects = Folder 'f-projects' '\\工作📮\Projects' 0 @($nested)
    $mailBelowCalendar = Folder 'f-calendar-mail' '\\工作📮\Calendar\Mail' 0 @()
    $calendar = Folder 'f-calendar' '\\工作📮\Calendar' 1 @($mailBelowCalendar)
    $deleted = Folder 'f-deleted' '\\工作📮\Deleted Items' 0 @()
    $junk = Folder 'f-junk' '\\工作📮\Junk' 0 @()
    $root = Folder 'f-root' '\\工作📮' 0 @($projects, $calendar, $deleted, $junk)
    $archiveRoot = Folder 'f-archive' '\\Archive' 0 @()
    $first = Node 'store-a' 'store' @{ StoreID = 's-a'; DisplayName = '工作📮' } @{ GetRootFolder = $root }
    $second = Node 'store-b' 'store' @{ StoreID = 's-b'; DisplayName = $(if ($DuplicateNames) { '工作📮' } else { 'Archive' }) } @{ GetRootFolder = $archiveRoot }
    $values = @($first, $second)
    if ($Empty) { $values = @() }
    $stores = Node 'stores' 'collection' @{ Count = $values.Count } @{ Item = $values }
    $session = Node 'session' 'session' @{ Stores = $stores; DefaultStore = $(if ($Empty) { $null } else { $first }); Class = 1 } @{}
    $first.methods.GetDefaultFolder = $root; $second.methods.GetDefaultFolder = $null
    $folderMap = @{}
    function Populate($folder, $store, $parent) {
        if ($folderMap.ContainsKey($folder.refId)) { return }
        $folderMap[$folder.refId] = $folder
        $folder.properties.StoreID = $store.properties.StoreID; $folder.properties.Store = $store; $folder.properties.Parent = $parent
        foreach ($child in $folder.properties.Folders.methods.Item) { Populate $child $store $folder }
    }
    Populate $root $first $session; Populate $archiveRoot $second $session
    $app = Node 'application' 'application' @{ Version = '16.0' } @{ GetNamespace = $session }
    $state.application = $app
    $state.folderMap = $folderMap
    $state.root = $root
    $state.archiveRoot = $archiveRoot
    $adapter = @{
        Connect = {
            $state.connections++
            [void]$state.calls.Add('Connect')
            if ($state.failures.ContainsKey('Connect')) { throw 'Synthetic COM/profile denied.' }
            [void]$state.acquired.Add($state.application.refId)
            @{ Application = $state.application; Mode = $state.mode }
        }.GetNewClosure()
        Property = {
            param($Reference, $Name)
            $key = $Reference.refId + '.' + $Name
            [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic property denied.' }
            if (-not $Reference.properties.ContainsKey($Name)) { throw 'Forbidden/unexpected property.' }
            $value = $Reference.properties[$Name]
            if ($value -is [MailLens.Tests.Reference]) { if ($null -ne $value) { [void]$state.acquired.Add($value.refId) } }
            return ,$value
        }.GetNewClosure()
        Method = {
            param($Reference, $Name, $Arguments)
            $key = $Reference.refId + '.' + $Name
            if ($Name -eq 'Item') { $key += '.' + $Arguments[0] }
            [void]$state.calls.Add($key)
            if ($state.failures.ContainsKey($key)) { throw 'Synthetic method denied.' }
            if ($Name -eq 'GetFolderFromID') {
                $value = $state.folderMap[$Arguments[0]]
                if ($null -eq $value -or $value.properties.StoreID -cne $Arguments[1]) { throw 'Unknown exact folder.' }
                [void]$state.acquired.Add($value.refId); return ,$value
            }
            if (-not $Reference.methods.ContainsKey($Name)) { throw 'Forbidden/unexpected method.' }
            if ($Name -eq 'Item') {
                if ($Arguments[0] -lt 1) { throw 'COM collections are 1-based.' }
                $value = $Reference.methods.Item[$Arguments[0] - 1]
            } else { $value = $Reference.methods[$Name] }
            if ($null -ne $value) { [void]$state.acquired.Add($value.refId) }
            return ,$value
        }.GetNewClosure()
        Release = {
            param($Reference)
            [void]$state.calls.Add($Reference.refId + '.Release')
            [void]$state.released.Add($Reference.refId)
        }.GetNewClosure()
    }
    @{ adapter = $adapter; state = $state }
}

function Assert-AllReleased($Double) {
    $acquired = @($Double.state.acquired.ToArray() | Sort-Object)
    $released = @($Double.state.released.ToArray() | Sort-Object)
    ($acquired -join '|') | Should -BeExactly ($released -join '|')
}
