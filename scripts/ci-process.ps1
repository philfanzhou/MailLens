# Development-only subprocess helpers. Paths remain data, never command text.
function Invoke-CIProcess {
    param([string]$Executable, [string[]]$Arguments, [string]$Directory, [int]$TimeoutSeconds = 90)
    if (-not [IO.File]::Exists($Executable)) { throw "Required executable missing: $Executable" }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    $quoted = foreach ($argument in $Arguments) {
        if ($argument.Contains('"') -or $argument.Contains("`n") -or $argument.Contains("`r")) { throw 'Unsupported subprocess argument.' }
        '"' + [regex]::Replace($argument, '(\\+)$', '$1$1') + '"'
    }
    $info.Arguments = $quoted -join ' '
    $info.WorkingDirectory = $Directory
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = New-Object Text.UTF8Encoding($false, $true)
    $info.StandardErrorEncoding = New-Object Text.UTF8Encoding($false, $true)
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw 'Subprocess failed to start.' }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill(); $process.WaitForExit()
            throw 'Owned CI subprocess timed out.'
        }
        return [pscustomobject]@{ Code = $process.ExitCode; Out = $stdout.GetAwaiter().GetResult(); Error = $stderr.GetAwaiter().GetResult() }
    } finally { $process.Dispose() }
}
function ConvertFrom-CISkills {
    param([string]$Json)
    if (-not $Json.TrimStart().StartsWith('[')) { throw 'CLI discovery did not return a JSON array.' }
    $records = ConvertFrom-Json -InputObject $Json -ErrorAction Stop
    foreach ($record in @($records)) {
        foreach ($field in @('name', 'description', 'location', 'content')) {
            if ($record.PSObject.Properties.Name -cnotcontains $field -or $record.$field -isnot [string]) { throw 'Invalid CLI skill record.' }
        }
    }
    return @($records)
}
function Assert-CITarget {
    param([object[]]$Records, [string]$SkillPath, [string]$Content, [string]$Description)
    $targets = @($Records | Where-Object { $_.name -ceq 'mail-lens' })
    if ($targets.Count -ne 1) { throw 'Expected exactly one mail-lens discovery record.' }
    $actual = [IO.Path]::GetFullPath($targets[0].location)
    if (-not $actual.Equals([IO.Path]::GetFullPath($SkillPath), [StringComparison]::OrdinalIgnoreCase)) { throw 'Wrong discovered target path.' }
    if ($targets[0].content.Trim() -cne $Content.Trim() -or $targets[0].description -cne $Description) { throw 'Wrong discovered target content/description.' }
}
