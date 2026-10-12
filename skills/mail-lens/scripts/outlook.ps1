# Run in native Windows PowerShell 5.1: powershell.exe -NoProfile -STA -File outlook.ps1 ...
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RequestPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
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
    $code = Invoke-MailLensFileRequest -RequestPath $RequestPath -OutputPath $OutputPath -CancellationRequested { [MailLens.ConsoleCancellation]::Requested }
} catch {
    # Fixed English message only: never print request contents, COM objects or mail.
    [Console]::Error.WriteLine('MailLens could not publish a result. Check absolute request/output paths and use a unique writable result file.')
} finally { [MailLens.ConsoleCancellation]::Stop() }
exit $code
