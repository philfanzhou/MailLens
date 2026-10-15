BeforeDiscovery { Import-Module (Join-Path $PSScriptRoot "../../skills/mail-lens/scripts/MailLens.Outlook.psm1") -Force }
BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/MailLens.Outlook.psm1') -Force
    . (Join-Path $PSScriptRoot '../fixtures/com/SessionDouble.ps1')
    . (Join-Path $PSScriptRoot '../fixtures/com/SearchDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/BrowseDouble.ps1')
    . (Join-Path $PSScriptRoot '../fixtures/com/DetailDouble.ps1')
}

Describe 'Managed caller and bound private selection' {
    InModuleScope MailLens.Outlook {
        BeforeAll {
            . (Join-Path $PSScriptRoot '../fixtures/com/SessionDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/SearchDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/BrowseDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/DetailDouble.ps1')
            $native=@{isWindows=$true;powerShellVersion='5.1.26100.1';edition='Desktop';apartmentState='STA';timeZone='UTC'}
            $utf8=New-Object Text.UTF8Encoding($false,$true)
            function New-CallerFixture {
                $root=Join-Path ([IO.Path]::GetTempPath()) ('caller-test-'+[Guid]::NewGuid().ToString('N'))
                [void][IO.Directory]::CreateDirectory($root)
                $input=Join-Path $root 'browse.json'
                [IO.File]::WriteAllText($input,'{"schemaVersion":2,"operation":"browse"}',(New-Object Text.UTF8Encoding($true)))
                return @{root=$root;request=$input;capture=@{json=$null};runs=(New-Object 'System.Collections.Generic.List[string]')}
            }
            function Invoke-CallerFixture($Fixture,$Double,$Extra=@{}) {
                $capture=$Fixture.capture
                $writer={param($Json) $capture.json=$Json}.GetNewClosure()
                $arguments=@{RequestPath=$Fixture.request;Adapter=$Double.adapter;Runtime=$native;WriteView=$writer}
                foreach($key in $Extra.Keys) {$arguments[$key]=$Extra[$key]}
                $code=Invoke-MailLensManagedRequest @arguments
                $view=ConvertFrom-Json $capture.json
                if ($view.runDirectory) {[void]$Fixture.runs.Add($view.runDirectory)}
                return @{code=$code;view=$view;json=$capture.json}
            }
            function Remove-CallerFixture($Fixture) {
                foreach ($run in $Fixture.runs) {if ([IO.Directory]::Exists($run)) {[IO.Directory]::Delete($run,$true)}}
                if ([IO.Directory]::Exists($Fixture.root)) {[IO.Directory]::Delete($Fixture.root,$true)}
            }
        }
        BeforeEach { $fixture=New-CallerFixture; $double=New-DetailDouble }
        AfterEach { Remove-CallerFixture $fixture }

        It 'executes in process, commits the bound index and repeats exact get in new runs' {
            $selection=Invoke-CallerFixture $fixture $double
            $selection.code | Should -BeIn @(0,2)
            $selection.view.summaries[0].label | Should -BeExactly M1
            $selection.json | Should -Not -Match '"(storeId|folderId|entryId|warnings|folderCoverage|preview|index)"'
            $run=$selection.view.runDirectory
            $rawBefore=Get-MailLensPrivateHash (Join-Path $run 'result.json')
            $indexBefore=Get-MailLensPrivateHash (Join-Path $run 'index.json')
            $index=Read-MailLensPrivateJson (Join-Path $run 'index.json')
            $index.entries[0].entryId | Should -BeExactly m-1
            $index.resultSha256 | Should -BeExactly $rawBefore
            $capture=$fixture.capture; $writer={param($Json) $capture.json=$Json}.GetNewClosure()
            $bodyRuns=@()
            foreach($iteration in 1..2) {
                $code=Invoke-MailLensManagedRequest -RunDirectory $run -Label M1 -Adapter $double.adapter -Runtime $native -WriteView $writer
                $code | Should -Be 0
                $body=ConvertFrom-Json $capture.json
                [void]$fixture.runs.Add($body.runDirectory); $bodyRuns+=$body.runDirectory
                $body.label | Should -BeExactly M1
                $body.selectionRunDirectory | Should -BeExactly $run
                $body.detail.body | Should -BeExactly "合同📮`r`nline"
                $body.detail.bodyLength | Should -Be 2000
                $body.detail.source.PSObject.Properties.Name | Should -Not -Contain entryId
            }
            $bodyRuns[0] | Should -Not -Be $bodyRuns[1]
            (Get-MailLensPrivateHash (Join-Path $run 'result.json')) | Should -BeExactly $rawBefore
            (Get-MailLensPrivateHash (Join-Path $run 'index.json')) | Should -BeExactly $indexBefore
            $double.state.calls | Should -Contain 'session.GetItemFromID.s-a.m-1'
            Assert-AllReleased $double
        }
        It 'commits an empty usable index for zero results' {
            $double=New-BrowseDouble -Count 0
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 0
            $result.view.selectedCount | Should -Be 0
            (Read-MailLensPrivateJson (Join-Path $result.view.runDirectory 'index.json')).entries.Count | Should -Be 0
        }
        It 'publishes only ten of twelve complete summaries with accurate omissions' {
            $double=New-BrowseDouble -Count 12
            [IO.File]::WriteAllText($fixture.request,'{"schemaVersion":2,"operation":"browse","limit":12}',$utf8)
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 0
            $result.view.selectedCount | Should -Be 10
            $result.view.omittedSummaryCount | Should -Be 2
            $result.view.metrics.returnedCount | Should -Be 12
            $result.view.summaries[9].label | Should -BeExactly M10
            @($double.state.calls | Where-Object {$_ -match '\.(Body|Recipients|Attachments|SenderName)$'}).Count | Should -Be 0
        }
        It 'preserves malformed UTF-8 request failure without any COM' {
            [IO.File]::WriteAllBytes($fixture.request,[byte[]]@(255))
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 1
            $result.view.errors[0].code | Should -BeExactly invalid_request_file
            $result.view.operation | Should -BeNullOrEmpty
            @($double.state.calls).Count | Should -Be 0
        }
        It 'does not authorize labels for failed or cancelled queries: <Kind>' -ForEach @(@{Kind='failed'},@{Kind='cancelled'}) {
            if($Kind -eq 'failed') {$double=New-DetailDouble -Failures @{'session.DefaultStore'=$true}; $result=Invoke-CallerFixture $fixture $double}
            else { $calls=@{n=0}; $cancel={$calls.n++;return $calls.n -gt 2}.GetNewClosure(); $result=Invoke-CallerFixture $fixture $double @{CancellationRequested=$cancel} }
            $result.code | Should -BeIn @(1,130)
            $result.view.selectedCount | Should -Be 0
            if($Kind -eq 'cancelled') {$result.view.errors[0].code | Should -BeExactly cancelled}
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
        }
        It 'preserves parse and request-validation failure operation null: <Text>' -ForEach @(@{Text='{'},@{Text='{"schemaVersion":1,"operation":"browse"}'},@{Text='{"schemaVersion":2,"operation":"browse","limit":"bad"}'}) {
            [IO.File]::WriteAllText($fixture.request,$Text,$utf8)
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 1
            $result.view.operation | Should -BeNullOrEmpty
            $result.view.errors[0].code | Should -Not -Be caller_unavailable
            @($double.state.calls).Count | Should -Be 0
        }
        It 'rejects stale or incompatible selections before any COM: <Mutation>' -ForEach @(
            @{Mutation='missing-index'},@{Mutation='invalid-json'},@{Mutation='version'},@{Mutation='moved'},@{Mutation='filename'},@{Mutation='hash'},@{Mutation='identity'},@{Mutation='position'},@{Mutation='label'},@{Mutation='deleted-result'},@{Mutation='invalid-utf8'},@{Mutation='failed-status'},@{Mutation='case-label'},@{Mutation='unknown-label'}
        ) {
            $selection=Invoke-CallerFixture $fixture $double; $run=$selection.view.runDirectory
            $indexPath=Join-Path $run 'index.json'; $rawPath=Join-Path $run 'result.json'; $index=Read-MailLensPrivateJson $indexPath; $label='M1'
            switch($Mutation) {
                missing-index {[IO.File]::Delete($indexPath)}
                invalid-json {[IO.File]::WriteAllText($indexPath,'{',$utf8)}
                version {$index.formatVersion=2}
                moved {$index.runDirectory=$fixture.root}
                filename {$index.resultFile='../result.json'}
                hash {[IO.File]::AppendAllText($rawPath,' ',$utf8)}
                identity {$index.entries[0].entryId='wrong-id'}
                position {$index.entries[0].summaryPosition=2}
                label {$index.entries[0].label='M2'}
                deleted-result {[IO.File]::Delete($rawPath)}
                invalid-utf8 {[IO.File]::WriteAllBytes($indexPath,[byte[]]@(255))}
                failed-status {$index.status='failed'}
                case-label {$label='m1'}
                unknown-label {$label='M10'}
            }
            if($Mutation -in @('version','moved','filename','identity','position','label','failed-status')) {[IO.File]::WriteAllText($indexPath,(ConvertTo-Json $index -Depth 20 -Compress),$utf8)}
            $before=$double.state.calls.Count; $capture=$fixture.capture
            $writer={param($Json)$capture.json=$Json}.GetNewClosure()
            (Invoke-MailLensManagedRequest -RunDirectory $run -Label $label -Adapter $double.adapter -Runtime $native -WriteView $writer) | Should -Be 1
            $double.state.calls.Count | Should -Be $before
            $capture.json | Should -Match 'caller_unavailable'
            $capture.json | Should -Not -Match 'wrong-id|s-a|m-1'
        }
        It 'rejects mixed modes and invalid ranges before COM: <Case>' -ForEach @(
            @{Case='mixed';ModeArguments=@{RequestPath='x';RunDirectory='x';Label='M1'}},@{Case='missing';ModeArguments=@{}},@{Case='offset';ModeArguments=@{RunDirectory='x';Label='M1';BodyOffset=-1}},@{Case='length';ModeArguments=@{RunDirectory='x';Label='M1';BodyLength=100001}},@{Case='query-body';ModeArguments=@{RequestPath='x';BodyLength=2000}},@{Case='relative';ModeArguments=@{RequestPath='relative.json'}},@{Case='empty-label';ModeArguments=@{RequestPath='x';Label=''}}
        ) {
            $capture=$fixture.capture; $writer={param($Json)$capture.json=$Json}.GetNewClosure()
            (Invoke-MailLensManagedRequest @ModeArguments -Adapter $double.adapter -Runtime $native -WriteView $writer) | Should -Be 1
            @($double.state.calls).Count | Should -Be 0
            $capture.json | Should -Match caller_unavailable
        }
        It 'retains exact-ID failure or body-change diagnostics through label get: <Kind>' -ForEach @(@{Kind='deleted'},@{Kind='changed'}) {
            $selection=Invoke-CallerFixture $fixture $double; $run=$selection.view.runDirectory
            $original=Get-MailLensPrivateHash (Join-Path $run 'index.json')
            if($Kind -eq 'deleted') {$double.state.lookup.Remove('s-a|m-1')}
            else {$double.state.actions['m-1.Body']={param($Item)$Item.properties.LastModificationTime=$Item.properties.LastModificationTime.AddSeconds(1)}}
            $capture=$fixture.capture;$writer={param($Json)$capture.json=$Json}.GetNewClosure()
            $code=Invoke-MailLensManagedRequest -RunDirectory $run -Label M1 -Adapter $double.adapter -Runtime $native -WriteView $writer
            $body=ConvertFrom-Json $capture.json; [void]$fixture.runs.Add($body.runDirectory)
            if($Kind -eq 'deleted') {$code | Should -Be 1; $body.PSObject.Properties.Name | Should -Not -Contain detail; $body.errors[0].code | Should -BeExactly item_unavailable}
            else {$code | Should -Be 2; $body.detail.contentChangedDuringRead | Should -BeTrue; $body.detail.consistencyAvailable | Should -BeFalse}
            (Get-MailLensPrivateHash (Join-Path $run 'index.json')) | Should -BeExactly $original
        }
        It 'rejects missing, malformed and mismatched raw publication without an index: <Kind>' -ForEach @(@{Kind='missing'},@{Kind='malformed'},@{Kind='mismatch'}) {
            Mock Invoke-MailLensFileRequest {
                if($Kind -eq 'malformed') {[IO.File]::WriteAllText($OutputPath,'{',$utf8)}
                if($Kind -eq 'mismatch') {
                    $raw=Invoke-MailLensRequest ([pscustomobject]@{schemaVersion=2;operation='browse'}) $Adapter $Runtime
                    [IO.File]::WriteAllText($OutputPath,(ConvertTo-Json $raw -Depth 30 -Compress),$utf8)
                }
                return 130
            }
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 1
            $result.json | Should -Match caller_unavailable
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
        }
        It 'rejects diagnose <RawStatus> with missing data through managed delivery' -ForEach @(@{RawStatus='complete'},@{RawStatus='partial'}) {
            [IO.File]::WriteAllText($fixture.request,'{"schemaVersion":2,"operation":"diagnose"}',$utf8)
            Mock Invoke-MailLensFileRequest {
                $raw=Invoke-MailLensRequest ([pscustomobject]@{schemaVersion=2;operation='diagnose'}) $Adapter $Runtime
                $raw.status=$RawStatus; $raw.data=$null
                [IO.File]::WriteAllText($OutputPath,(ConvertTo-Json $raw -Depth 30 -Compress),$utf8)
                return (Get-MailLensExitCode $RawStatus)
            }
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 1
            $result.view.status | Should -BeExactly failed
            $result.view.exitCode | Should -Be 1
            $result.view.errors[0].code | Should -BeExactly caller_unavailable
            $result.view.errors[0].message | Should -BeExactly 'MailLens caller state or output is unavailable; coverage is unknown.'
            $result.view.operation | Should -BeNullOrEmpty
            $result.json | Should -Not -Match '"(summaries|detail|body|label|coverage)"'
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
            (Read-MailLensPrivateJson (Join-Path $result.view.runDirectory 'result.json')).status | Should -BeExactly $RawStatus
        }
        It 'preserves valid managed diagnose when version metadata is <VersionState>' -ForEach @(@{VersionState='available';ExpectedCode=0},@{VersionState='unavailable';ExpectedCode=2}) {
            [IO.File]::WriteAllText($fixture.request,'{"schemaVersion":2,"operation":"diagnose"}',$utf8)
            if($VersionState -eq 'unavailable') {$double=New-DetailDouble -Failures @{'application.Version'=$true}}
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be $ExpectedCode
            $result.view.operation | Should -BeExactly diagnose
            $result.view.coverage.evaluated | Should -BeFalse
            $result.json | Should -Not -Match '"(summaries|detail|body|label)"|caller_unavailable'
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
            Assert-AllReleased $double
        }
        It 'keeps raw success private when index publication fails' {
            Mock Write-MailLensPrivateJson {throw 'Synthetic secret identity'} -ParameterFilter {$Path.EndsWith('index.json')}
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 1
            $result.json | Should -Not -Match 'Synthetic|identity|summaries|body'
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'result.json')) | Should -BeTrue
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
        }
        It 'cancels before index commit without publishing success' {
            Mock Write-MailLensPrivateJson {Stop-MailLens cancelled 'Cancelled.'} -ParameterFilter {$Path.EndsWith('index.json')}
            $result=Invoke-CallerFixture $fixture $double
            $result.code | Should -Be 130
            $result.json | Should -Not -Match 'summaries|body'
            [IO.File]::Exists((Join-Path $result.view.runDirectory 'index.json')) | Should -BeFalse
        }
        It 'reports stdout delivery failure and leaves committed private state intact' {
            $saved=@{run=$null;writes=0}
            $writer={param($Json)$saved.writes++; $value=ConvertFrom-Json $Json; if($value.status -in @('complete','partial')) {$saved.run=$value.runDirectory};throw 'Synthetic stdout denied'}.GetNewClosure()
            (Invoke-MailLensManagedRequest -RequestPath $fixture.request -Adapter $double.adapter -Runtime $native -WriteView $writer) | Should -Be 1
            [void]$fixture.runs.Add($saved.run)
            [IO.File]::Exists((Join-Path $saved.run 'index.json')) | Should -BeTrue
        }
        It 'preserves existing output and reservation on publication collisions' {
            $path=Join-Path $fixture.root 'index.json'; [IO.File]::WriteAllText($path,'preserve',$utf8)
            {Write-MailLensPrivateJson $path @{x=1}} | Should -Throw
            [IO.File]::ReadAllText($path) | Should -BeExactly preserve
            @(Get-ChildItem $fixture.root -Filter '.maillens-*.tmp').Count | Should -Be 0
        }
        It 'rejects reparse-point selection files before COM' -Tag Native {
            $selection=Invoke-CallerFixture $fixture $double; $run=$selection.view.runDirectory; $index=Join-Path $run 'index.json'
            $target=Join-Path $fixture.root 'original-index.json'; [IO.File]::Move($index,$target)
            # A junction requires no Windows developer-mode or symlink privilege.
            $junction=Join-Path $fixture.root 'linked-run'
            New-Item -ItemType Junction -Path $junction -Target $run -ErrorAction Stop | Out-Null
            try {
                $before=$double.state.calls.Count; $capture=$fixture.capture; $writer={param($Json)$capture.json=$Json}.GetNewClosure()
                (Invoke-MailLensManagedRequest -RunDirectory $junction -Label M1 -Adapter $double.adapter -Runtime $native -WriteView $writer) | Should -Be 1
                $double.state.calls.Count | Should -Be $before
            } finally {[IO.Directory]::Delete($junction)}
        }
        It 'rejects repository request ancestry before COM' {
            $repo=Join-Path $fixture.root 'repo'; [void][IO.Directory]::CreateDirectory($repo); [IO.File]::WriteAllText((Join-Path $repo '.git'),'gitdir: synthetic',$utf8)
            $fixture.request=Join-Path $repo 'request.json'; [IO.File]::WriteAllText($fixture.request,'{}',$utf8)
            (Invoke-CallerFixture $fixture $double).code | Should -Be 1
            @($double.state.calls).Count | Should -Be 0
        }
    }
}

Describe 'Strict compact projection validation and allowlist' {
    InModuleScope MailLens.Outlook {
        BeforeAll {
            . (Join-Path $PSScriptRoot '../fixtures/com/SessionDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/SearchDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/BrowseDouble.ps1')
            . (Join-Path $PSScriptRoot '../fixtures/com/DetailDouble.ps1')
            $native=@{isWindows=$true;powerShellVersion='5.1.26100.1';edition='Desktop';apartmentState='STA';timeZone='UTC'}
            function ProjectionFixture($Operation='browse') {
                $request=[pscustomobject]@{schemaVersion=2;operation=$Operation}
                if($Operation -eq 'search') {$request|Add-Member listOnly $true}
                $double=New-DetailDouble
                $raw=Invoke-MailLensRequest $request $double.adapter $native
                $result=ConvertFrom-Json (ConvertTo-Json $raw -Depth 30 -Compress)
                return @{request=$request;result=$result;double=$double}
            }
        }
        It 'rejects malformed envelopes, data and selected identities: <Mutation>' -ForEach @(
            @{Mutation='schema'},@{Mutation='status'},@{Mutation='operation'},@{Mutation='exit'},@{Mutation='missing-data'},@{Mutation='missing-summaries'},@{Mutation='identity-type'},@{Mutation='empty-identity'},@{Mutation='count-type'},@{Mutation='subject-type'},@{Mutation='warning-count'},@{Mutation='missing-error'}
        ) {
            $f=ProjectionFixture; $result=$f.result; $code=Get-MailLensExitCode $result.status
            switch($Mutation) {
                schema {$result.schemaVersion=3}
                status {$result.status='unknown'}
                operation {$result.operation='search'}
                exit {$code=130}
                missing-data {$result.data=$null}
                missing-summaries {$result.data.browse.PSObject.Properties.Remove('summaries')}
                identity-type {$result.data.browse.summaries[0].entryId=7}
                empty-identity {$result.data.browse.summaries[0].storeId=''}
                count-type {$result.data.browse.visitedCount='1'}
                subject-type {$result.data.browse.summaries[0].subject=@{unsafe='raw'}}
                warning-count {$result.warningCount=999}
                missing-error {$result.status='failed';$result.errors=@();$code=1}
            }
            {ConvertTo-MailLensCompactView $result $code $f.request 'C:\synthetic'} | Should -Throw
        }
        It 'limits labels to ten and preserves all effective metrics, gaps and warning totals' {
            $f=ProjectionFixture; $result=$f.result; $rows=@()
            foreach($i in 1..12) {$row=ConvertFrom-Json (ConvertTo-Json $result.data.browse.summaries[0]);$row.entryId='synthetic-'+$i;$rows+=$row}
            $result.data.browse.summaries=$rows;$result.data.browse.returnedCount=12
            $p=ConvertTo-MailLensCompactView $result (Get-MailLensExitCode $result.status) $f.request 'C:\synthetic'
            $p.entries.Count | Should -Be 10
            $p.view.selectedCount | Should -Be 10
            $p.view.omittedSummaryCount | Should -Be 2
            $p.entries[9].summaryPosition | Should -Be 10
            $p.view.warningCount | Should -Be $result.warningCount
            $p.view.metrics.visitedCount | Should -Be $result.data.browse.visitedCount
            (ConvertTo-Json $p.view -Depth 20) | Should -Not -Match 'synthetic-12|"entryId"|"warnings"'
        }
        It 'preserves search terms, range, budgets and aggregate per-folder gaps without rows' {
            $f=ProjectionFixture search
            $p=ConvertTo-MailLensCompactView $f.result (Get-MailLensExitCode $f.result.status) $f.request 'C:\synthetic'
            $p.view.range.timeZone | Should -BeExactly UTC
            $p.view.listOnly | Should -BeTrue
            $p.view.metrics.summaryLimit | Should -Be 20
            $p.view.searchCoverage.attemptedRows | Should -Be $f.result.data.search.folderCoverage.Count
            (ConvertTo-Json $p.view -Depth 20) | Should -Not -Match '"folderCoverage"|"storeId"|"folderId"'
        }
        It 'supports diagnose and folders aggregate views without an index' -ForEach @(@{Operation='diagnose'},@{Operation='folders'}) {
            $f=ProjectionFixture $Operation
            $p=ConvertTo-MailLensCompactView $f.result (Get-MailLensExitCode $f.result.status) $f.request 'C:\synthetic'
            $p.entries.Count | Should -Be 0
            $p.view.Contains('summaries') | Should -BeFalse
            $p.view.coverage.evaluated | Should -Be ($Operation -eq 'folders')
        }
    }
}

Describe 'Native redirected stdout and concurrent label reads' -Tag Native {
    BeforeAll {
        . (Join-Path $PSScriptRoot '../../scripts/ci-process.ps1')
        $fixtureScript=Join-Path $PSScriptRoot '../fixtures/com/CallerProcess.ps1'
        $launcher=Join-Path $PSScriptRoot '../../skills/mail-lens/scripts/outlook.ps1'
        $powershell=Join-Path $PSHOME 'powershell.exe'
        $utf8=New-Object Text.UTF8Encoding($false,$true)
    }
    BeforeEach {
        $root=Join-Path ([IO.Path]::GetTempPath()) ('caller-process-'+[Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($root); $runs=@();$jobs=@()
        $request=Join-Path $root 'request.json';[IO.File]::WriteAllText($request,'{"schemaVersion":2,"operation":"browse"}',$utf8)
    }
    AfterEach {
        foreach($job in $jobs) {Stop-Job $job -ErrorAction SilentlyContinue;Remove-Job $job -Force -ErrorAction SilentlyContinue}
        foreach($run in $runs) {if([IO.Directory]::Exists($run)) {[IO.Directory]::Delete($run,$true)}}
        if([IO.Directory]::Exists($root)) {[IO.Directory]::Delete($root,$true)}
    }
    It 'writes one BOM-free UTF-8 object under an OEM 437 console and ASCII OutputEncoding' {
        $info=New-Object Diagnostics.ProcessStartInfo
        $info.FileName=$powershell; $info.Arguments='-NoProfile -STA -File "'+$fixtureScript+'" -RequestPath "'+$request+'"'
        $info.WorkingDirectory=$root; $info.UseShellExecute=$false; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
        $process=New-Object Diagnostics.Process; $process.StartInfo=$info; $bytes=New-Object IO.MemoryStream
        try {
            [void]$process.Start()
            $copy=$process.StandardOutput.BaseStream.CopyToAsync($bytes); $stderr=$process.StandardError.ReadToEndAsync()
            if(-not $process.WaitForExit(30000)) {$process.Kill(); throw 'Synthetic child timed out.'}
            $copy.GetAwaiter().GetResult(); $null=$stderr.GetAwaiter().GetResult()
            $rawBytes=$bytes.ToArray(); $rawBytes[0] | Should -Be 123
            ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 239 -and $rawBytes[1] -eq 187 -and $rawBytes[2] -eq 191) | Should -BeFalse
            $result=@{Code=$process.ExitCode;Out=$utf8.GetString($rawBytes)}
        } finally {if(-not $process.HasExited) {$process.Kill()}; $process.Dispose(); $bytes.Dispose()}
        $result.Code | Should -BeIn @(0,2)
        $result.Out[0] | Should -Be '{'
        $view=ConvertFrom-Json $result.Out;$runs+=$view.runDirectory
        $view.summaries[0].subject | Should -BeExactly '合成📮 stdout'
        $view.resolvedScope.stores[0].storeName | Should -BeExactly '工作📮'
        $result.Out | Should -Not -Match '"storeId"|"entryId"|"warnings"|"preview"'
        @($result.Out.Trim() -split "`n").Count | Should -Be 1
    }
    It 'permits overlapping label subprocesses with distinct outputs and an unchanged source index' {
        $result=Invoke-CIProcess $powershell @('-NoProfile','-STA','-File',$fixtureScript,'-RequestPath',$request) $root
        $selection=ConvertFrom-Json $result.Out;$runs+=$selection.runDirectory
        $indexPath=Join-Path $selection.runDirectory 'index.json';$original=[IO.File]::ReadAllBytes($indexPath)
        foreach($i in 1..2) {
            $viewFile=Join-Path $root ('view-'+$i+'.json')
            $jobs+=Start-Job -ScriptBlock {param($Script,$Run,$Output) & $Script -RunDirectory $Run -Label M1 -ViewPath $Output; if($LASTEXITCODE -ne 0) {throw 'Synthetic child failed.'}} -ArgumentList $fixtureScript,$selection.runDirectory,$viewFile
        }
        $finished=@($jobs | Wait-Job -Timeout 45)
        $finished.Count | Should -Be 2
        foreach($job in $jobs) {Receive-Job $job -ErrorAction Stop; $job.State | Should -Be Completed}
        $first=ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $root 'view-1.json'),$utf8));$second=ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $root 'view-2.json'),$utf8))
        $runs+=$first.runDirectory;$runs+=$second.runDirectory
        $first.runDirectory | Should -Not -Be $second.runDirectory
        $first.detail.body | Should -BeExactly "合同📮`r`nline"
        $second.detail.body | Should -BeExactly $first.detail.body
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($indexPath)) | Should -BeExactly ([Convert]::ToBase64String($original))
    }
    It 'keeps raw launcher parse failures compatible and rejects mixed launcher modes' {
        [IO.File]::WriteAllText($request,'{',$utf8);$output=Join-Path $root 'result.json'
        $raw=Invoke-CIProcess $powershell @('-NoProfile','-STA','-File',$launcher,'-RequestPath',$request,'-OutputPath',$output) $root
        $raw.Code | Should -Be 1
        $raw.Out | Should -BeNullOrEmpty
        (ConvertFrom-Json ([IO.File]::ReadAllText($output,$utf8))).errors[0].code | Should -BeExactly invalid_request_file
        $mixed=Invoke-CIProcess $powershell @('-NoProfile','-STA','-File',$launcher,'-View','-RequestPath',$request,'-OutputPath',(Join-Path $root 'must-not-exist.json')) $root
        $mixed.Code | Should -Be 1
        [IO.File]::Exists((Join-Path $root 'must-not-exist.json')) | Should -BeFalse
    }
}
