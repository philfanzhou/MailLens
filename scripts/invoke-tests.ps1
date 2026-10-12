[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestPath, [Parameter(Mandatory)][string]$PesterManifest,
    [string]$RequiredContainers = '', [string]$ReportPath)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
try {
    if (-not [IO.File]::Exists($PesterManifest)) { throw 'Pinned Pester command is missing.' }
    Import-Module $PesterManifest -Force
    if ((Get-Module Pester).Version -ne [version]'5.7.1' -or -not (Get-Command Invoke-Pester -ErrorAction Stop)) { throw 'Pester 5.7.1 is required.' }
    if (-not [IO.Directory]::Exists($TestPath)) { throw 'Mandatory test path is missing.' }
    $files = @(Get-ChildItem -LiteralPath $TestPath -Filter '*.Tests.ps1' -File -Recurse)
    if ($files.Count -eq 0) { throw 'No test containers found.' }
    foreach ($name in $RequiredContainers.Split(';') | Where-Object { $_ }) {
        if (@($files | Where-Object { $_.Name -ceq $name }).Count -ne 1) { throw "Mandatory test container missing or duplicated: $name" }
    }
    $config = New-PesterConfiguration
    $config.Run.Path = @($files.FullName)
    $config.Run.PassThru = $true
    $config.Output.Verbosity = 'Normal'
    if ($ReportPath) { $config.TestResult.Enabled = $true; $config.TestResult.OutputPath = $ReportPath; $config.TestResult.OutputFormat = 'NUnitXml' }
    $result = Invoke-Pester -Configuration $config
    Write-Output "Pester $((Get-Module Pester).Version): containers=$($result.Containers.Count); total=$($result.TotalCount); passed=$($result.PassedCount); failed=$($result.FailedCount); skipped=$($result.SkippedCount); inconclusive=$($result.InconclusiveCount); not-run=$($result.NotRunCount)."
    if ($result.Result -ne 'Passed' -or $result.TotalCount -le 0 -or $result.PassedCount -ne $result.TotalCount -or $result.FailedCount -ne 0 -or $result.SkippedCount -ne 0 -or $result.InconclusiveCount -ne 0 -or $result.NotRunCount -ne 0 -or $result.Containers.Count -ne $files.Count) { throw 'Mandatory test coverage did not pass completely.' }
    foreach ($container in $result.Containers) {
        if ($container.Result -ne 'Passed' -or $container.TotalCount -le 0) { throw 'Empty or unsuccessful test container.' }
    }
    exit 0
} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
