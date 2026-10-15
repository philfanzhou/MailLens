# Synthetic subprocess fixture: never loads actual Outlook COM or live mail.
[CmdletBinding()]
param([string]$RequestPath, [string]$RunDirectory, [string]$Label, [string]$ViewPath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::GetEncoding(437)
$OutputEncoding=[Text.Encoding]::ASCII
Import-Module (Join-Path $PSScriptRoot '../../../skills/mail-lens/scripts/MailLens.Outlook.psm1') -Force
. (Join-Path $PSScriptRoot 'SessionDouble.ps1')
. (Join-Path $PSScriptRoot 'SearchDouble.ps1')
. (Join-Path $PSScriptRoot 'DetailDouble.ps1')
$double=New-DetailDouble
$double.state.mail.properties.Subject='合成📮 stdout'
$native=@{isWindows=$true;powerShellVersion='5.1.26100.1';edition='Desktop';apartmentState='STA';timeZone='UTC'}
$arguments=@{Adapter=$double.adapter;Runtime=$native}
if($RunDirectory) {$arguments.RunDirectory=$RunDirectory;$arguments.Label=$Label} else {$arguments.RequestPath=$RequestPath}
if($ViewPath) {
    $arguments.WriteView={param($Json) [IO.File]::WriteAllText($ViewPath,$Json,(New-Object Text.UTF8Encoding($false,$true)))}.GetNewClosure()
}
$code=Invoke-MailLensManagedRequest @arguments
exit $code
