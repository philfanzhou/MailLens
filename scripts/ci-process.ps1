# Development-only subprocess helpers. Paths remain data, never command text.
function Invoke-CIProcess {
    param([string]$Executable, [string[]]$Arguments, [string]$Directory, [int]$TimeoutSeconds = 90, [hashtable]$Environment = @{})
    if (-not [IO.File]::Exists($Executable)) { throw "Required executable missing: $Executable" }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    $quoted = foreach ($argument in $Arguments) {
        if ($argument.Contains('"') -or $argument.Contains("`n") -or $argument.Contains("`r")) { throw 'Unsupported subprocess argument.' }
        '"' + [regex]::Replace($argument, '(\\+)$', '$1$1') + '"'
    }
    foreach ($key in $Environment.Keys) {
        if ($null -eq $Environment[$key]) { $info.EnvironmentVariables.Remove($key) } else { $info.EnvironmentVariables[$key] = $Environment[$key] }
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
    $started = $false
    try {
        if (-not $process.Start()) { throw 'Subprocess failed to start.' }
        $started = $true
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill(); [void]$process.WaitForExit(5000)
            throw 'Owned CI subprocess timed out.'
        }
        return [pscustomobject]@{ Code = $process.ExitCode; Out = $stdout.GetAwaiter().GetResult(); Error = $stderr.GetAwaiter().GetResult() }
    } finally {
        if ($started -and -not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
        $process.Dispose()
    }
}
# The V2 endpoints return location envelopes; arrays alone are the obsolete V1 surface.
function ConvertFrom-CIEnvelope {
    param([string]$Json, [string]$Directory)
    try { $value = ConvertFrom-Json -InputObject $Json.TrimStart([char]0xfeff) -ErrorAction Stop } catch { throw 'Invalid V2 JSON response.' }
    if ($null -eq $value -or $value -isnot [pscustomobject] -or
        $value.PSObject.Properties.Name -cnotcontains 'location' -or
        $value.location -isnot [pscustomobject] -or
        $value.location.PSObject.Properties.Name -cnotcontains 'directory' -or
        $value.location.directory -isnot [string] -or
        -not [IO.Path]::IsPathRooted($value.location.directory) -or
        $value.PSObject.Properties.Name -cnotcontains 'data' -or $value.data -isnot [array]) { throw 'Invalid V2 location/data envelope.' }
    if ($Directory -and -not ([IO.Path]::GetFullPath($value.location.directory)).Equals([IO.Path]::GetFullPath($Directory), [StringComparison]::OrdinalIgnoreCase)) { throw 'Wrong V2 response location.' }
    return $value
}
function ConvertFrom-CISkills {
    param([string]$Json, [string]$Directory)
    $value = ConvertFrom-CIEnvelope $Json $Directory
    foreach ($record in $value.data) {
        if ($record -isnot [pscustomobject]) { throw 'Invalid V2 skill record.' }
        foreach ($field in @('id', 'name', 'path', 'content')) {
            if ($record.PSObject.Properties.Name -cnotcontains $field -or $record.$field -isnot [string]) { throw 'Invalid V2 skill record.' }
        }
        if (-not [IO.Path]::IsPathRooted($record.path)) { throw 'Invalid V2 absolute skill path.' }
        if ($record.PSObject.Properties.Name -ccontains 'description' -and $record.description -isnot [string]) { throw 'Invalid V2 skill description.' }
        if ($record.PSObject.Properties.Name -ccontains 'autoinvoke' -and $record.autoinvoke -isnot [bool]) { throw 'Invalid V2 skill autoinvoke.' }
    }
    return $value.data
}
function Assert-CITarget {
    param([object[]]$Records, [string]$SkillPath, [string]$Content, [string]$Description)
    $targets = @($Records | Where-Object { $_.id -ceq 'mail-lens' })
    if ($targets.Count -ne 1) { throw 'Expected exactly one mail-lens discovery id.' }
    $target = $targets[0]
    if ($target.name -cne 'mail-lens') { throw 'Wrong discovered target name.' }
    if (-not [IO.Path]::IsPathRooted($target.path) -or -not ([IO.Path]::GetFullPath($target.path)).Equals([IO.Path]::GetFullPath($SkillPath), [StringComparison]::OrdinalIgnoreCase)) { throw 'Wrong discovered target path.' }
    if ($target.PSObject.Properties.Name -cnotcontains 'description' -or $target.content -cne $Content -or $target.description -cne $Description) { throw 'Wrong discovered target content/description.' }
}
function Test-CIPluginReady {
    param([string]$Json, [string]$Directory)
    $value = ConvertFrom-CIEnvelope $Json $Directory
    $matches = @()
    foreach ($record in $value.data) {
        if ($record -isnot [pscustomobject] -or $record.PSObject.Properties.Name -cnotcontains 'state' -or
            $record.state -isnot [pscustomobject] -or $record.state.PSObject.Properties.Name -cnotcontains 'status' -or
            $record.state.status -cnotin @('active', 'failed')) { throw 'Invalid V2 plugin state.' }
        if ($record.PSObject.Properties.Name -ccontains 'id') {
            if ($record.id -isnot [string]) { throw 'Invalid V2 plugin id.' }
            if ($record.id -ceq 'opencode.config.skill') { $matches += $record }
        }
    }
    if ($matches.Count -gt 1) { throw 'Duplicate V2 discovery plugin.' }
    if ($matches.Count -eq 0) { return $false }
    if ($matches[0].state.status -ceq 'failed') { throw 'V2 discovery plugin activation failed.' }
    return $true
}
function Get-CIAvailablePort {
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    try { $listener.Start(); return $listener.LocalEndpoint.Port } finally { $listener.Stop() }
}
function Start-CIService {
    param([string]$Executable, [string]$Directory, [int]$Port = 0)
    if (-not [IO.File]::Exists($Executable)) { throw 'Required service executable missing.' }
    if ($Port -eq 0) { $Port = Get-CIAvailablePort }
    # A port bind race fails closed; no response from a pre-existing listener is adopted.
    $reservation = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $Port)
    try { $reservation.Start() } catch { throw 'Owned service port is occupied.' } finally { $reservation.Stop() }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    $info.Arguments = 'serve --hostname 127.0.0.1 --port ' + $Port
    $info.WorkingDirectory = $Directory
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = New-Object Text.UTF8Encoding($false, $true)
    $info.StandardErrorEncoding = New-Object Text.UTF8Encoding($false, $true)
    $secret = [Guid]::NewGuid().ToString('N') + [Guid]::NewGuid().ToString('N')
    $info.EnvironmentVariables['OPENCODE_SERVER_PASSWORD'] = $secret
    $info.EnvironmentVariables['OPENCODE_PASSWORD'] = $secret
    $info.EnvironmentVariables['OPENCODE_SERVER_USERNAME'] = 'opencode'
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    $started = $false
    try {
        if (-not $process.Start()) { throw 'Owned service failed to start.' }
        $started = $true
        # Discard both streams after draining: generated authentication and service internals never enter CI logs.
        $out = $process.StandardOutput.BaseStream.CopyToAsync([IO.Stream]::Null)
        $err = $process.StandardError.BaseStream.CopyToAsync([IO.Stream]::Null)
        return [pscustomobject]@{ Process=$process; DrainOut=$out; DrainError=$err; Port=$Port; Directory=$Directory
            Base=('http://127.0.0.1:' + $Port); Authorization=('Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('opencode:' + $secret))) }
    } catch {
        if ($started -and -not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
        $process.Dispose(); throw 'Owned service failed to start.'
    }
}
function Stop-CIService {
    param($Service)
    if ($null -eq $Service) { return }
    try {
        if (-not $Service.Process.HasExited) { $Service.Process.Kill() }
        if (-not $Service.Process.WaitForExit(5000)) { throw 'Owned service cleanup timed out.' }
        if (-not $Service.DrainOut.Wait(5000) -or -not $Service.DrainError.Wait(5000)) { throw 'Owned service stream cleanup timed out.' }
    } finally { $Service.Process.Dispose(); $Service.Authorization = $null }
}
function Invoke-CIServiceRequest {
    param($Service, [string]$Path, [string]$Method = 'GET', [string]$Body, [int]$TimeoutMilliseconds = 5000, [switch]$Initializing)
    if ($Service.Process.HasExited) { throw 'Owned service exited.' }
    $request = [Net.HttpWebRequest]::Create($Service.Base + $Path)
    $request.Proxy = $null; $request.AllowAutoRedirect = $false
    $request.Method = $Method; $request.Timeout = $TimeoutMilliseconds; $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Headers['Authorization'] = $Service.Authorization
    $response = $null; $stream = $null
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($Method -eq 'POST') {
            $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
            $request.ContentType = 'application/json'; $request.ContentLength = $bytes.Length
            $stream = $request.GetRequestStream(); $stream.Write($bytes, 0, $bytes.Length); $stream.Dispose(); $stream = $null
        }
        $response = $request.GetResponse()
        if ([int]$response.StatusCode -lt 200 -or [int]$response.StatusCode -ge 300) { throw 'Owned service HTTP status failure.' }
        $stream = $response.GetResponseStream()
        $buffer = New-Object byte[] 8192; $memory = New-Object IO.MemoryStream
        try {
            while (($count = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                if ($memory.Length + $count -gt 4194304 -or $clock.ElapsedMilliseconds -gt $TimeoutMilliseconds) { throw 'Owned service response exceeded bounds.' }
                $memory.Write($buffer, 0, $count)
            }
            $utf8 = New-Object Text.UTF8Encoding($false, $true)
            return $utf8.GetString($memory.ToArray()).TrimStart([char]0xfeff)
        } finally { $memory.Dispose() }
    } catch [Net.WebException] {
        # Connection refusal is transient only while this owned process starts listening.
        $refused = $_.Exception.Status -eq [Net.WebExceptionStatus]::ConnectFailure
        $inner = $_.Exception.InnerException
        while ($null -ne $inner) {
            if ($inner -is [Net.Sockets.SocketException] -and $inner.SocketErrorCode -eq [Net.Sockets.SocketError]::ConnectionRefused) { $refused = $true }
            $inner = $inner.InnerException
        }
        if ($Initializing -and $refused) { return $null }
        throw ('Owned service HTTP/authentication/transport failure: ' + $_.Exception.Status + '.')
    } catch { throw 'Owned service request/UTF-8 response failed.' }
    finally { if ($null -ne $stream) { $stream.Dispose() }; if ($null -ne $response) { $response.Dispose() }; $request.Abort() }
}
function Wait-CISkills {
    param($Service, [int]$TimeoutSeconds = 30)
    $clock = [Diagnostics.Stopwatch]::StartNew(); $attempts = 0
    $query = '?location%5Bdirectory%5D=' + [Uri]::EscapeDataString($Service.Directory)
    while ($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ($Service.Process.HasExited) { throw 'Owned service exited before discovery readiness.' }
        $remaining = [Math]::Max(1, [Math]::Min(3000, $TimeoutSeconds * 1000 - [int]$clock.ElapsedMilliseconds))
        $plugins = Invoke-CIServiceRequest $Service ('/api/plugin' + $query) -TimeoutMilliseconds $remaining -Initializing
        $attempts++
        if ($null -ne $plugins -and (Test-CIPluginReady $plugins $Service.Directory)) {
            Write-Host "V2 skill plugin active after $attempts observation(s)."
            $remaining = $TimeoutSeconds * 1000 - [int]$clock.ElapsedMilliseconds
            if ($remaining -le 0) { throw 'Owned service discovery readiness timed out.' }
            $remaining = [Math]::Min(3000, $remaining)
            $skills = Invoke-CIServiceRequest $Service ('/api/skill' + $query) -TimeoutMilliseconds $remaining
            return ConvertFrom-CISkills $skills $Service.Directory
        }
        Start-Sleep -Milliseconds 100
    }
    throw 'Owned service discovery readiness timed out.'
}
function Assert-CISkillActivation {
    param($Service, [string]$Content)
    $sessionID = $null
    try {
        $body = @{ location=@{directory=$Service.Directory}; title='MailLens synthetic skill admission' } | ConvertTo-Json -Compress
        $created = ConvertFrom-Json (Invoke-CIServiceRequest $Service '/api/session' 'POST' $body)
        $sessionID = $created.data.id
        if ($sessionID -isnot [string] -or $sessionID -notmatch '^ses_[a-zA-Z0-9]+$') { throw 'Invalid owned session id.' }
        [void](Invoke-CIServiceRequest $Service ('/api/experimental/session/' + $sessionID + '/skill') 'POST' '{"id":"mail-lens","resume":false}')
        $messages = ConvertFrom-Json (Invoke-CIServiceRequest $Service ('/api/session/' + $sessionID + '/context'))
        $skills = @($messages.data | Where-Object { $_.type -ceq 'skill' })
        if ($skills.Count -ne 1 -or $skills[0].skill -cne 'mail-lens' -or $skills[0].name -cne 'mail-lens' -or $skills[0].text -cne $Content) { throw 'Wrong stored V2 skill activation body.' }
        # Stored skill messages are body-only. The model tool wrapper is separately defined by pinned Skill.prepare/toModelOutput.
        Write-Host 'Explicit resume:false skill body admission passed; no model execution requested.'
    } catch { throw 'Owned V2 skill activation failed.' } finally {
        if ($null -ne $sessionID -and $sessionID -match '^ses_[a-zA-Z0-9]+$' -and -not $Service.Process.HasExited) {
            [void](Invoke-CIServiceRequest $Service ('/api/session/' + $sessionID) 'DELETE')
        }
    }
}

function Expand-CIOpenCodeArchive {
    param([string]$Archive, [string]$Destination, [string]$Integrity)
    $algorithm = [Security.Cryptography.SHA512]::Create(); $inputFile = [IO.File]::OpenRead($Archive)
    try { $actual = [Convert]::ToBase64String($algorithm.ComputeHash($inputFile)) } finally { $algorithm.Dispose(); $inputFile.Dispose() }
    if ($actual -cne $Integrity) { throw 'OpenCode archive digest mismatch.' }
    if ([IO.Directory]::Exists($Destination)) { throw 'OpenCode tool directory already exists.' }
    [void][IO.Directory]::CreateDirectory($Destination)
    $inputFile = [IO.File]::OpenRead($Archive)
    $gzip = New-Object IO.Compression.GzipStream($inputFile, [IO.Compression.CompressionMode]::Decompress)
    $seen = @{}; $outputPath = Join-Path $Destination 'opencode.exe'; $output = $null
    function Read-TarBlock([int]$Length) {
        $bytes = New-Object byte[] $Length; $offset = 0
        while ($offset -lt $Length) { $read = $gzip.Read($bytes, $offset, $Length - $offset); if ($read -eq 0) { throw 'Truncated OpenCode archive.' }; $offset += $read }
        return ,$bytes
    }
    try {
        while ($true) {
            $header = Read-TarBlock 512
            if (@($header | Where-Object { $_ -ne 0 }).Count -eq 0) {
                # Two zero records end the tar. Refuse hidden entries after that terminator.
                $tail = Read-TarBlock 512
                if (@($tail | Where-Object { $_ -ne 0 }).Count -ne 0) { throw 'Invalid OpenCode archive terminator.' }
                $buffer = New-Object byte[] 8192
                while (($count = $gzip.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    for ($i=0; $i -lt $count; $i++) { if ($buffer[$i] -ne 0) { throw 'Unexpected trailing OpenCode archive data.' } }
                }
                break
            }
            $checksum = [Text.Encoding]::ASCII.GetString($header,148,8).Trim([char]0,' ')
            if ($checksum -notmatch '^[0-7]+$') { throw 'Invalid OpenCode tar checksum.' }
            $sum=0; for ($i=0; $i -lt 512; $i++) { if ($i -ge 148 -and $i -lt 156) { $sum+=32 } else { $sum += [int]$header[$i] } }
            if ($sum -ne [Convert]::ToInt64($checksum,8)) { throw 'OpenCode tar checksum mismatch.' }
            $name = [Text.Encoding]::ASCII.GetString($header,0,100).TrimEnd([char]0)
            $prefix = [Text.Encoding]::ASCII.GetString($header,345,155).TrimEnd([char]0)
            if ($prefix -or $name -cnotin @('package/package.json','package/bin/opencode.exe') -or $seen.ContainsKey($name) -or $header[156] -notin @(0,48)) { throw 'Unsafe or ambiguous OpenCode archive entry.' }
            $seen[$name]=$true
            $sizeText = [Text.Encoding]::ASCII.GetString($header,124,12).Trim([char]0,' ')
            if ($sizeText -notmatch '^[0-7]+$') { throw 'Invalid OpenCode archive size.' }
            $size = [Convert]::ToInt64($sizeText,8)
            if ($size -le 0 -or $size -gt 314572800) { throw 'OpenCode archive entry exceeds bounds.' }
            if ($name -ceq 'package/bin/opencode.exe') { $output = [IO.File]::Open($outputPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None) }
            $remaining=$size
            while ($remaining -gt 0) { $block=Read-TarBlock ([int][Math]::Min(65536,$remaining)); if ($null -ne $output) { $output.Write($block,0,$block.Length) }; $remaining-=$block.Length }
            if ($null -ne $output) { $output.Dispose(); $output=$null }
            $padding = (512 - $size % 512) % 512
            if ($padding -gt 0) { [void](Read-TarBlock ([int]$padding)) }
        }
        if ($seen.Count -ne 2 -or -not $seen.ContainsKey('package/bin/opencode.exe')) { throw 'Pinned OpenCode executable missing or ambiguous.' }
        return $outputPath
    } catch {
        if ($null -ne $output) { $output.Dispose(); $output=$null }
        if ([IO.File]::Exists($outputPath)) { [IO.File]::Delete($outputPath) }
        throw
    } finally { if ($null -ne $output) { $output.Dispose() }; $gzip.Dispose(); $inputFile.Dispose() }
}
