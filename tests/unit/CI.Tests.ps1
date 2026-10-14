BeforeAll {
    $script:Repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    . (Join-Path $script:Repository 'scripts/check-package.ps1')
    . (Join-Path $script:Repository 'scripts/ci-process.ps1')
    $script:Utf8 = New-Object Text.UTF8Encoding($false, $true)
    $script:PowerShell = (Get-Command powershell.exe -CommandType Application -ErrorAction Stop).Source
    function Write-CIText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:Utf8) }
    function Invoke-NegativeTests([string]$Path, [string]$Manifest = $env:MAIL_LENS_PESTER_MANIFEST, [string]$Required = '') {
        Invoke-CIProcess $script:PowerShell @('-NoProfile', '-STA', '-File', (Join-Path $script:Repository 'scripts/invoke-tests.ps1'), '-TestPath', $Path, '-PesterManifest', $Manifest, '-RequiredContainers', $Required) $script:CaseRoot
    }
}
Describe 'CI failure propagation' {
    BeforeEach {
        $script:CaseRoot = Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-ci-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($script:CaseRoot)
        $script:Copy = Join-Path $script:CaseRoot 'source'
        [void][IO.Directory]::CreateDirectory($script:Copy)
        foreach ($name in @('skills', 'scripts', 'VERSION', 'LICENSE')) { Copy-Item -LiteralPath (Join-Path $script:Repository $name) -Destination $script:Copy -Recurse }
    }
    AfterEach { if ([IO.Directory]::Exists($script:CaseRoot)) { [IO.Directory]::Delete($script:CaseRoot, $true) } }
    It 'accepts the actual package' { { Test-MailLensPackage $script:Copy } | Should -Not -Throw }
    It 'rejects broken package <Fault>' -ForEach @(
        @{Fault='missing-entry'}, @{Fault='missing-frontmatter'}, @{Fault='empty-description'},
        @{Fault='duplicate-field'}, @{Fault='wrong-name'}, @{Fault='bad-field'}, @{Fault='missing-resource'},
        @{Fault='escaping-resource'}, @{Fault='missing-version'}, @{Fault='bad-version'}, @{Fault='duplicate-version'},
        @{Fault='missing-bom'}, @{Fault='malformed-utf8'}, @{Fault='parse-error'}, @{Fault='whitespace'}, @{Fault='comment-language'}) {
        $entry = Join-Path $script:Copy 'skills/mail-lens/SKILL.md'
        $text = [IO.File]::ReadAllText($entry, $script:Utf8)
        switch ($Fault) {
            'missing-entry' { [IO.File]::Delete($entry) }
            'missing-frontmatter' { Write-CIText $entry "# Synthetic skill`n" }
            'empty-description' { Write-CIText $entry "---`nname: mail-lens`ndescription: `n---`n" }
            'duplicate-field' { Write-CIText $entry "---`nname: mail-lens`nname: mail-lens`ndescription: synthetic`n---`n" }
            'wrong-name' { Write-CIText $entry "---`nname: wrong`ndescription: synthetic`n---`n" }
            'bad-field' { Write-CIText $entry "---`nname: mail-lens`ndescription: [array]`n---`n" }
            'missing-resource' { [IO.File]::Delete((Join-Path $script:Copy 'skills/mail-lens/references/workflow.md')) }
            'escaping-resource' { Write-CIText $entry ($text + "`n[escape](../../VERSION)`n") }
            'missing-version' { [IO.File]::Delete((Join-Path $script:Copy 'VERSION')) }
            'bad-version' { Write-CIText (Join-Path $script:Copy 'VERSION') "latest`n" }
            'duplicate-version' { Write-CIText $entry ($text.Replace('name: mail-lens', "name: mail-lens`nversion: 0.1.0")) }
            'missing-bom' { Write-CIText (Join-Path $script:Copy 'scripts/install.ps1') "# synthetic`n" }
            'malformed-utf8' { [IO.File]::WriteAllBytes($entry, [byte[]]@(255,254,255)) }
            'parse-error' { [IO.File]::WriteAllText((Join-Path $script:Copy 'scripts/install.ps1'), "function {`n", (New-Object Text.UTF8Encoding($true, $true))) }
            'whitespace' { Write-CIText $entry ($text + "bad `n") }
            'comment-language' { [IO.File]::WriteAllText((Join-Path $script:Copy 'scripts/install.ps1'), ('# ' + [char]0x4e2d + "`n"), (New-Object Text.UTF8Encoding($true, $true))) }
        }
        { Test-MailLensPackage $script:Copy } | Should -Throw
    }
    It 'rejects missing command/path, empty containers and required-container deletion' -ForEach @(@{Fault='command'}, @{Fault='path'}, @{Fault='empty'}, @{Fault='required'}) {
        $testPath = Join-Path $script:CaseRoot 'tests'
        [void][IO.Directory]::CreateDirectory($testPath)
        $manifest = $env:MAIL_LENS_PESTER_MANIFEST
        $required = ''
        switch ($Fault) {
            'command' { $manifest = Join-Path $script:CaseRoot 'missing.psd1' }
            'path' { $testPath = Join-Path $script:CaseRoot 'missing' }
            'empty' { Write-CIText (Join-Path $testPath 'Empty.Tests.ps1') "Describe 'empty' { }`n" }
            'required' { Write-CIText (Join-Path $testPath 'Pass.Tests.ps1') "Describe 'pass' { It 'works' { 1 | Should -Be 1 } }`n"; $required = 'Missing.Tests.ps1' }
        }
        (Invoke-NegativeTests $testPath $manifest $required).Code | Should -Be 1
    }
    It 'propagates deliberately failing and skipped synthetic tests' -ForEach @(@{Kind='failure'}, @{Kind='skip'}) {
        $testPath = Join-Path $script:CaseRoot 'tests'; [void][IO.Directory]::CreateDirectory($testPath)
        if ($Kind -eq 'failure') { $test = "Describe 'negative' { It 'fails' { 1 | Should -Be 2 } }`n" }
        else { $test = "Describe 'negative' { It 'skips' -Skip { 1 | Should -Be 1 } }`n" }
        Write-CIText (Join-Path $testPath 'Negative.Tests.ps1') $test
        (Invoke-NegativeTests $testPath).Code | Should -Be 1
    }
    It 'rejects absent CLI executable and times out only its owned child' {
        { Invoke-CIProcess (Join-Path $script:CaseRoot 'absent.exe') @() $script:CaseRoot } | Should -Throw '*missing*'
        $scriptPath = Join-Path $script:CaseRoot 'sleep.ps1'; Write-CIText $scriptPath "Start-Sleep -Seconds 30`n"
        { Invoke-CIProcess $script:PowerShell @('-NoProfile', '-File', $scriptPath) $script:CaseRoot 1 } | Should -Throw '*timed out*'
    }

}

Describe 'V2 discovery protocol and owned service failures' {
    BeforeAll {
        function New-CIEnvelope($Data, [string]$Directory = $script:CaseRoot) { @{location=@{directory=$Directory};data=@($Data)} | ConvertTo-Json -Depth 12 -Compress }
        function New-CIRecord { [pscustomobject]@{id='mail-lens';name='mail-lens';path=(Join-Path $script:CaseRoot 'SKILL.md');description='Synthetic Chinese description';content='Complete synthetic body'} }
        function Start-CIFakeServer([int]$Status = 200, [int]$Delay = 0) {
            $port=Get-CIAvailablePort
            $file=Join-Path $script:CaseRoot 'server.ps1'; $ready=Join-Path $script:CaseRoot 'ready'
            Write-CIText $file @'
param([int]$Port,[string]$Ready,[int]$Status,[int]$Delay)
$listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$Port)
try {
    $listener.Start(); [IO.File]::WriteAllText($Ready,'ready')
    $client=$listener.AcceptTcpClient(); $stream=$client.GetStream()
    $buffer=New-Object byte[] 4096; [void]$stream.Read($buffer,0,$buffer.Length)
    if ($Delay -gt 0) { Start-Sleep -Seconds $Delay }
    $bytes=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $Status Synthetic`r`nContent-Length: 2`r`nConnection: close`r`n`r`n{}")
    $stream.Write($bytes,0,$bytes.Length); $stream.Dispose(); $client.Close()
    Start-Sleep -Seconds 30
} finally { $listener.Stop() }
'@
            $info=New-Object Diagnostics.ProcessStartInfo
            $info.FileName=$script:PowerShell; $info.Arguments='-NoProfile -File "' + $file + '" -Port ' + $port + ' -Ready "' + $ready + '" -Status ' + $Status + ' -Delay ' + $Delay
            $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
            $process=New-Object Diagnostics.Process; $process.StartInfo=$info; [void]$process.Start()
            $script:Service=[pscustomobject]@{Process=$process;Port=$port;Base=('http://127.0.0.1:'+$port);Directory=$script:CaseRoot;Authorization='Basic synthetic-secret-must-never-appear';DrainOut=$process.StandardOutput.BaseStream.CopyToAsync([IO.Stream]::Null);DrainError=$process.StandardError.BaseStream.CopyToAsync([IO.Stream]::Null)}
            $clock=[Diagnostics.Stopwatch]::StartNew()
            while (-not [IO.File]::Exists($ready)) { if ($process.HasExited -or $clock.Elapsed.TotalSeconds -gt 10) { throw 'Synthetic HTTP server did not start.' }; Start-Sleep -Milliseconds 50 }
            return $script:Service
        }
    }
    BeforeEach {
        $script:CaseRoot=Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-v2-unit-'+[Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($script:CaseRoot); $script:Service=$null
    }
    AfterEach { try { Stop-CIService $script:Service } finally { if ([IO.Directory]::Exists($script:CaseRoot)) { [IO.Directory]::Delete($script:CaseRoot,$true) } } }
    It 'accepts a BOM, optional unrelated descriptions, Unicode and exact target fields' {
        $record=New-CIRecord; $record.description=[string][char]0x4e2d; $record.content=([string][char]0x6587)+' synthetic body'
        $other=[pscustomobject]@{id='builtin';name='built in';path=(Join-Path $script:CaseRoot 'builtin.md');content='body'}
        $records=@(ConvertFrom-CISkills ([string][char]0xfeff+(New-CIEnvelope @($record,$other))) $script:CaseRoot)
        $records.Count | Should -Be 2
        { Assert-CITarget $records $record.path $record.content $record.description } | Should -Not -Throw
    }
    It 'rejects invalid V2 envelope or record <Fault>' -ForEach @(
        @{Fault='v1'},@{Fault='null'},@{Fault='array-data'},@{Fault='location'},@{Fault='other-location'},@{Fault='record'},@{Fault='id'},@{Fault='path'},@{Fault='relative'},@{Fault='content'},@{Fault='description'},@{Fault='autoinvoke'}) {
        $record=New-CIRecord; $json=New-CIEnvelope @($record)
        switch($Fault) {
            'v1' { $json='[]' }; 'null' { $json='null' }; 'array-data' { $json='{"location":{"directory":"C:/synthetic"},"data":{}}' }
            'location' { $json='{"location":null,"data":[]}' }; 'other-location' { $json=New-CIEnvelope @($record) (Join-Path $script:CaseRoot 'other') }
            'record' { $json=New-CIEnvelope @('bad') }; 'id' { $record.id=42; $json=New-CIEnvelope @($record) }
            'path' { $record.path=$null; $json=New-CIEnvelope @($record) }; 'relative' { $record.path='relative.md'; $json=New-CIEnvelope @($record) }
            'content' { $record.content=@('not string'); $json=New-CIEnvelope @($record) }
            'description' { $record.description=42; $json=New-CIEnvelope @($record) }
            'autoinvoke' { $record | Add-Member autoinvoke 'yes'; $json=New-CIEnvelope @($record) }
        }
        { ConvertFrom-CISkills $json $script:CaseRoot } | Should -Throw
    }
    It 'rejects a wrong or duplicate target <Fault>' -ForEach @(@{Fault='missing'},@{Fault='wrong-id'},@{Fault='duplicate'},@{Fault='name'},@{Fault='path'},@{Fault='body'},@{Fault='description'},@{Fault='body-whitespace'}) {
        $record=New-CIRecord; $path=$record.path; $body=$record.content; $description=$record.description; $records=@($record)
        switch($Fault) {
            'missing' {$records=@()}; 'wrong-id' {$record.id='MAIL-LENS'}; 'duplicate' {$records=@($record,$record)}
            'name' {$record.name='other'}; 'path' {$record.path=Join-Path $script:CaseRoot 'other.md'}
            'body' {$record.content='wrong'}; 'description' {$record.description='wrong'}; 'body-whitespace' {$record.content+=' '}
        }
        { Assert-CITarget $records $path $body $description } | Should -Throw
    }
    It 'waits for plugin activation rather than accepting an early empty list' {
        $script:observations=0
        Mock Invoke-CIServiceRequest {
            if ($Path -like '/api/plugin*') {
                $script:observations++
                if ($script:observations -eq 1) { return New-CIEnvelope @() }
                return New-CIEnvelope @(@{id='opencode.config.skill';state=@{status='active'}})
            }
            return New-CIEnvelope @((New-CIRecord))
        }
        $records=@(Wait-CISkills ([pscustomobject]@{Process=@{HasExited=$false};Directory=$script:CaseRoot}) 2)
        $records.Count | Should -Be 1; $script:observations | Should -Be 2
        Should -Invoke Invoke-CIServiceRequest -Times 1 -Exactly -ParameterFilter { $Path -like '/api/skill*' }
    }
    It 'accepts true ready absence but fails a positive assertion' {
        Mock Invoke-CIServiceRequest {
            if ($Path -like '/api/plugin*') { return New-CIEnvelope @(@{id='opencode.config.skill';state=@{status='active'}}) }
            New-CIEnvelope @()
        }
        $records=@(Wait-CISkills ([pscustomobject]@{Process=@{HasExited=$false};Directory=$script:CaseRoot}) 2)
        $records.Count | Should -Be 0
        { Assert-CITarget $records (Join-Path $script:CaseRoot 'SKILL.md') 'body' 'description' } | Should -Throw '*exactly one*'
    }
    It 'rejects failed/malformed plugin readiness <Fault>' -ForEach @(@{Fault='failed'},@{Fault='malformed'},@{Fault='duplicate'}) {
        $records=@(@{id='opencode.config.skill';state=@{status='active'}})
        switch($Fault) { 'failed' {$records[0].state.status='failed'}; 'malformed' {$records[0].state.status='initializing'}; 'duplicate' {$records+=$records[0]} }
        { Test-CIPluginReady (New-CIEnvelope $records) $script:CaseRoot } | Should -Throw
    }
    It 'bounds initialization and detects a dead owned service' {
        Mock Invoke-CIServiceRequest { New-CIEnvelope @() }
        $service=[pscustomobject]@{Process=@{HasExited=$false};Directory=$script:CaseRoot}
        { Wait-CISkills $service 1 } | Should -Throw '*timed out*'
        $service.Process.HasExited=$true
        { Wait-CISkills $service 1 } | Should -Throw '*exited*'
    }
    It 'refuses an occupied port without affecting its owner' {
        $listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0)
        try { $listener.Start(); $port=$listener.LocalEndpoint.Port; { Start-CIService $script:PowerShell $script:CaseRoot $port } | Should -Throw '*occupied*'; $listener.Server.IsBound | Should -BeTrue } finally { $listener.Stop() }
    }
    It 'fails HTTP/authentication errors with no authentication in exceptions <Status>' -ForEach @(@{Status=401},@{Status=503}) {
        $service=Start-CIFakeServer $Status
        $failure=$null
        try { Invoke-CIServiceRequest $service '/api/plugin' } catch {$failure=$_.Exception.Message}
        $failure | Should -BeLike '*HTTP/authentication/transport failure*'
        $failure | Should -Not -Match 'synthetic-secret'
    }
    It 'cleans the owned service when the caller cooperatively cancels' {
        $service=Start-CIFakeServer; $ownedID=$service.Process.Id
        try {
            try { throw (New-Object OperationCanceledException('Synthetic cancellation.')) }
            finally { Stop-CIService $service; $script:Service=$null }
        } catch [OperationCanceledException] { }
        @(Get-Process -Id $ownedID -ErrorAction SilentlyContinue).Count | Should -Be 0
        $service.Authorization | Should -BeNullOrEmpty
    }
    It 'bounds HTTP waits and cleans its owned child on assertion failure' {
        $service=Start-CIFakeServer 200 30; $ownedID=$service.Process.Id
        try { { Invoke-CIServiceRequest $service '/api/plugin' -TimeoutMilliseconds 100 } | Should -Throw '*failure*' }
        finally { Stop-CIService $service; $script:Service=$null }
        @(Get-Process -Id $ownedID -ErrorAction SilentlyContinue).Count | Should -Be 0
        $listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$service.Port)
        try { { $listener.Start() } | Should -Not -Throw } finally { $listener.Stop() }
    }
}

Describe 'Integrity-pinned executable extraction' {
    BeforeAll {
        function New-CITestArchive([string]$Fault) {
            $tar=New-Object IO.MemoryStream
            try {
                $names=@('package/package.json','package/bin/opencode.exe')
                switch($Fault) {
                    'traversal' {$names[1]='../opencode.exe'}; 'absolute' {$names[1]='/opencode.exe'}
                    'duplicate' {$names+=$names[1]}; 'missing' {$names=@($names[0])}
                }
                foreach($name in $names) {
                    $header=New-Object byte[] 512; $content=[Text.Encoding]::ASCII.GetBytes('synthetic executable bytes')
                    [Text.Encoding]::ASCII.GetBytes($name).CopyTo($header,0)
                    [Text.Encoding]::ASCII.GetBytes(([Convert]::ToString($content.Length,8).PadLeft(11,'0')+[char]0)).CopyTo($header,124)
                    $header[156]=48
                    if($Fault -eq 'link' -and $name -eq $names[1]){$header[156]=50}
                    for($i=148;$i -lt 156;$i++){$header[$i]=32}
                    $sum=0;foreach($b in $header){$sum+=[int]$b}
                    [Text.Encoding]::ASCII.GetBytes(([Convert]::ToString($sum,8).PadLeft(6,'0')+[char]0+' ')).CopyTo($header,148)
                    if($Fault -eq 'checksum'){$header[0]=88}
                    $tar.Write($header,0,512);$tar.Write($content,0,$content.Length)
                    $padding=New-Object byte[] (512-$content.Length);$tar.Write($padding,0,$padding.Length)
                }
                $end=New-Object byte[] 1024;$tar.Write($end,0,$end.Length)
                if($Fault -eq 'trailing'){$tar.WriteByte(1)}
                $path=Join-Path $script:CaseRoot 'fixture.tgz';$file=[IO.File]::Create($path)
                $gzip=New-Object IO.Compression.GzipStream($file,[IO.Compression.CompressionMode]::Compress)
                try{$bytes=$tar.ToArray();if($Fault -eq 'truncated'){$bytes=$bytes[0..700]};$gzip.Write($bytes,0,$bytes.Length)}finally{$gzip.Dispose();$file.Dispose()}
                $algorithm=[Security.Cryptography.SHA512]::Create();try{$digest=[Convert]::ToBase64String($algorithm.ComputeHash([IO.File]::ReadAllBytes($path)))}finally{$algorithm.Dispose()}
                return [pscustomobject]@{Path=$path;Integrity=$digest}
            }finally{$tar.Dispose()}
        }
    }
    BeforeEach {$script:CaseRoot=Join-Path ([IO.Path]::GetTempPath()) ('mail-lens-archive-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($script:CaseRoot)}
    AfterEach {if([IO.Directory]::Exists($script:CaseRoot)){[IO.Directory]::Delete($script:CaseRoot,$true)}}
    It 'extracts only the single expected executable after verifying SHA-512' {
        $archive=New-CITestArchive ''; $target=Join-Path $script:CaseRoot 'tool'
        $path=Expand-CIOpenCodeArchive $archive.Path $target $archive.Integrity
        [IO.File]::ReadAllText($path) | Should -BeExactly 'synthetic executable bytes'
        @(Get-ChildItem $target -File -Recurse).Count | Should -Be 1
    }
    It 'rejects a wrong digest before creating an extraction directory' {
        $archive=New-CITestArchive ''; $target=Join-Path $script:CaseRoot 'tool'
        {Expand-CIOpenCodeArchive $archive.Path $target 'wrong'} | Should -Throw '*digest*'
        [IO.Directory]::Exists($target) | Should -BeFalse
    }
    It 'rejects unsafe, ambiguous, incomplete archives <Fault>' -ForEach @(@{Fault='traversal'},@{Fault='absolute'},@{Fault='duplicate'},@{Fault='missing'},@{Fault='link'},@{Fault='checksum'},@{Fault='trailing'},@{Fault='truncated'}) {
        $archive=New-CITestArchive $Fault;$target=Join-Path $script:CaseRoot 'tool'
        {Expand-CIOpenCodeArchive $archive.Path $target $archive.Integrity} | Should -Throw
        [IO.File]::Exists((Join-Path $target 'opencode.exe')) | Should -BeFalse
    }
}
