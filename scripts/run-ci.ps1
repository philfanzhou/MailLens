[CmdletBinding()]
param([Parameter(Mandatory)][ValidateSet('windows-tests', 'opencode-discovery')][string]$Job,
    [Parameter(Mandatory)][string]$RunRoot)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
try {
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1 -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'CI requires native Windows PowerShell 5.1 STA.' }
    $repository = Split-Path $PSScriptRoot -Parent
    $sha = & git -C $repository rev-parse HEAD
    if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { throw 'Cannot record checkout SHA.' }
    Write-Output "Checkout SHA: $sha; PowerShell $($PSVersionTable.PSVersion); apartment=$([Threading.Thread]::CurrentThread.ApartmentState)."
    if (Test-Path -LiteralPath $RunRoot) { throw 'CI run root already exists.' }
    [void][IO.Directory]::CreateDirectory($RunRoot)
    $env:TEMP = Join-Path $RunRoot 'temp'; $env:TMP = $env:TEMP
    [void][IO.Directory]::CreateDirectory($env:TEMP)
    . (Join-Path $PSScriptRoot 'check-package.ps1')
    Test-MailLensPackage $repository
    $modules = Join-Path $RunRoot 'modules'
    Save-Module Pester -RequiredVersion 5.7.1 -Repository PSGallery -Path $modules -Force -ErrorAction Stop
    $manifest = Join-Path $modules 'Pester/5.7.1/Pester.psd1'
    $env:MAIL_LENS_PESTER_MANIFEST = $manifest
    if ($Job -eq 'windows-tests') {
        $path = Join-Path $repository 'tests/unit'
        $required = 'Installer.Tests.ps1;Outlook.Browse.Tests.ps1;Outlook.Detail.Tests.ps1;Outlook.Search.Tests.ps1;Outlook.Session.Tests.ps1;CI.Tests.ps1'
    } else {
        $url = 'https://github.com/anomalyco/opencode/releases/download/v1.18.30/opencode-windows-x64-baseline.zip'
        $digest = '8c34c5767a688e8d6c4640b4893d34f08f08b071ed0043c940dc5be17a3b9233'
        $archive = Join-Path $RunRoot 'opencode.zip'
        Write-Output "OpenCode provenance: $url; release API SHA-256=$digest; source=5cd8e68fdd72b27818d26d168b9c7a06b359567e."
        Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $archive -TimeoutSec 180
        if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -cne $digest) { throw 'OpenCode archive digest mismatch.' }
        $tools = Join-Path $RunRoot 'opencode'
        Expand-Archive -LiteralPath $archive -DestinationPath $tools
        $executables = @(Get-ChildItem -LiteralPath $tools -Filter opencode.exe -File -Recurse)
        if ($executables.Count -ne 1) { throw 'Pinned OpenCode executable missing or ambiguous.' }
        $env:MAIL_LENS_OPENCODE = $executables[0].FullName
        $path = Join-Path $repository 'tests/integration'
        $required = 'OpenCode.Discovery.Tests.ps1'
    }
    & powershell.exe -NoProfile -STA -File (Join-Path $PSScriptRoot 'invoke-tests.ps1') -TestPath $path -PesterManifest $manifest -RequiredContainers $required -ReportPath (Join-Path $RunRoot 'synthetic-tests.xml')
    if ($LASTEXITCODE -ne 0) { throw 'Mandatory CI tests failed.' }
    exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
