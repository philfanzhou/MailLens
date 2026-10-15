# Run in native Windows PowerShell 5.1: powershell.exe -NoProfile -STA -File outlook.ps1 ...
[CmdletBinding()]
param(
    [string]$RequestPath,
    [string]$OutputPath,
    [switch]$View,
    [string]$RunDirectory,
    [string]$Label,
    [string]$BodyOffset = '0',
    [string]$BodyLength = '2000'
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'MailLens.Outlook.psm1') -Force
# A console Ctrl+C requests cancellation between COM calls. A blocked/killed host
# may have no final envelope; callers must treat a missing result as failure/unknown.
if (-not ('MailLens.ConsoleCancellation' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
namespace MailLens {
    public static class ConsoleCancellation {
        public static volatile bool Requested;
        private static void OnCancel(object sender, ConsoleCancelEventArgs args) {
            Requested = true; args.Cancel = true;
        }
        public static void Start() { Requested = false; Console.CancelKeyPress += OnCancel; }
        public static void Stop() { Console.CancelKeyPress -= OnCancel; }
    }
}
'@
}
$code = 1
try {
    [MailLens.ConsoleCancellation]::Start()
    if ($View) {
        if ($PSBoundParameters.ContainsKey('OutputPath')) { throw 'Invalid managed parameters.' }
        $arguments = @{ CancellationRequested = { [MailLens.ConsoleCancellation]::Requested } }
        foreach ($name in @('RequestPath', 'RunDirectory', 'Label', 'BodyOffset', 'BodyLength')) {
            if ($PSBoundParameters.ContainsKey($name)) {
                $value = $PSBoundParameters[$name]
                if ($name -in @('BodyOffset', 'BodyLength')) {
                    $number = 0
                    if ($value -notmatch '^[0-9]+$' -or -not [int]::TryParse($value, [ref]$number)) { throw 'Invalid body range.' }
                    $value = $number
                }
                $arguments[$name] = $value
            }
        }
        $code = Invoke-MailLensManagedRequest @arguments
    } else {
        foreach ($name in @('RunDirectory', 'Label', 'BodyOffset', 'BodyLength')) {
            if ($PSBoundParameters.ContainsKey($name)) { throw 'Invalid raw parameters.' }
        }
        if (-not $RequestPath -or -not $OutputPath) { throw 'Raw paths are required.' }
        $code = Invoke-MailLensFileRequest -RequestPath $RequestPath -OutputPath $OutputPath -CancellationRequested { [MailLens.ConsoleCancellation]::Requested }
    }
} catch {
    # Fixed English message only: never print request contents, COM objects or mail.
    [Console]::Error.WriteLine('MailLens could not publish a result. Check absolute request/output paths and use a unique writable result file.')
} finally { [MailLens.ConsoleCancellation]::Stop() }
exit $code
