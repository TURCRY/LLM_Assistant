param(
    [string]$JobsRoot = "C:\Affaires\_jobs",
    [string]$AsrScript = "D:\GPT4all_local\scripts\transcriptions\transcrire_local_voxtral.ps1",
    [string]$AnnotationBatch = "D:\GPT4all_local\scripts\projet_photos\run_all_photos_pcfixe.bat",
    [string]$ServerHostPort = "127.0.0.1:5050",
    [string]$DeepSeekEndpoint = "http://127.0.0.1:5050/ocr_deepseek_batch",
    [ValidateSet("fail", "requeue", "ignore")]
    [string]$RecoverRunning = "fail",
    [int]$StaleLockMinutes = 30,
    [switch]$Simulate
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$AllowedRoots = @(
    "C:\Affaires",
    "D:\GPT4all_local"
)

function Get-FullPathSafe {
    param([Parameter(Mandatory=$true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath($PathValue)
}

function Test-PathUnderRoot {
    param(
        [Parameter(Mandatory=$true)][string]$PathValue,
        [Parameter(Mandatory=$true)][string]$RootValue
    )
    $full = (Get-FullPathSafe -PathValue $PathValue).TrimEnd('\')
    $root = (Get-FullPathSafe -PathValue $RootValue).TrimEnd('\')
    return $full.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($root + "\", [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-AllowedPath {
    param(
        [Parameter(Mandatory=$true)][string]$PathValue,
        [Parameter(Mandatory=$true)][string]$FieldName
    )
    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        throw "Champ chemin vide: $FieldName"
    }
    foreach ($root in $AllowedRoots) {
        if (Test-PathUnderRoot -PathValue $PathValue -RootValue $root) {
            return Get-FullPathSafe -PathValue $PathValue
        }
    }
    throw "Chemin refuse pour ${FieldName}: $PathValue"
}

function Get-JobString {
    param(
        [Parameter(Mandatory=$true)][object]$Job,
        [Parameter(Mandatory=$true)][string]$Name
    )
    $prop = $Job.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return "" }
    return [string]$prop.Value
}

function Get-JobValue {
    param(
        [Parameter(Mandatory=$true)][object]$Job,
        [Parameter(Mandatory=$true)][string]$Name
    )
    $prop = $Job.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Assert-StrictBoolean {
    param(
        [Parameter(Mandatory=$true)][object]$Value,
        [Parameter(Mandatory=$true)][string]$FieldName
    )
    if ($Value -is [bool]) { return [bool]$Value }
    throw "Booleen strict attendu pour ${FieldName}"
}

function Assert-TileCount {
    param([Parameter(Mandatory=$true)][object]$Value)
    if ($Value -isnot [int] -and $Value -isnot [long]) {
        throw "tile_count doit etre un entier"
    }
    $tileCount = [int]$Value
    if (@(2, 3) -notcontains $tileCount) {
        throw "tile_count refuse: $tileCount"
    }
    return $tileCount
}

function Assert-Pages {
    param([Parameter(Mandatory=$true)][object]$Value)
    if ($Value -is [int] -or $Value -is [long]) {
        $Value = @($Value)
    }
    elseif ($Value -isnot [System.Array]) {
        throw "pages doit etre une liste d'entiers"
    }
    if ($Value.Count -lt 1) {
        throw "pages doit contenir au moins une page"
    }
    $pages = @()
    foreach ($page in $Value) {
        if ($page -isnot [int] -and $page -isnot [long]) {
            throw "pages doit contenir uniquement des entiers"
        }
        $pageInt = [int]$page
        if ($pageInt -lt 1 -or $pageInt -gt 10000) {
            throw "page refusee: $pageInt"
        }
        if ($pages -contains $pageInt) {
            throw "page dupliquee: $pageInt"
        }
        $pages += $pageInt
    }
    return ,$pages
}

function Quote-Arg {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Join-CommandLine {
    param([string[]]$ArgumentValues)
    return ($ArgumentValues | ForEach-Object { Quote-Arg $_ }) -join " "
}

function Quote-CmdArg {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Ensure-JobDirs {
    param([string]$Root)
    foreach ($name in @("queued", "running", "done", "failed", "logs")) {
        New-Item -ItemType Directory -Path (Join-Path $Root $name) -Force | Out-Null
    }
}

function Move-JobFile {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$DestinationDir
    )
    New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    $dest = Join-Path $DestinationDir (Split-Path $Source -Leaf)
    if (Test-Path -LiteralPath $dest) {
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $dest = Join-Path $DestinationDir ("{0}.{1}.json" -f ([IO.Path]::GetFileNameWithoutExtension($Source)), $stamp)
    }
    Move-Item -LiteralPath $Source -Destination $dest
    return $dest
}

function Read-LockPid {
    param([string]$LockPath)
    if (-not (Test-Path -LiteralPath $LockPath)) { return $null }
    try {
        $text = Get-Content -LiteralPath $LockPath -Raw -ErrorAction Stop
        if ($text -match 'pid=(\d+)') { return [int]$Matches[1] }
    }
    catch {
        return $null
    }
    return $null
}

function Open-SpoolerLock {
    param(
        [string]$LockPath,
        [int]$StaleMinutes
    )
    try {
        return [System.IO.File]::Open($LockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    }
    catch [System.IO.IOException] {
        $pidFromLock = Read-LockPid -LockPath $LockPath
        if ($pidFromLock) {
            $running = Get-Process -Id $pidFromLock -ErrorAction SilentlyContinue
            if ($running) {
                Write-Host "Spooler deja en cours: pid=$pidFromLock"
                return $null
            }
        }

        $item = Get-Item -LiteralPath $LockPath -ErrorAction SilentlyContinue
        if ($item -and $item.LastWriteTimeUtc -gt (Get-Date).ToUniversalTime().AddMinutes(-1 * $StaleMinutes)) {
            Write-Host "Verrou present et non perime: $LockPath"
            return $null
        }

        Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue
        return [System.IO.File]::Open($LockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    }
}

function Invoke-LoggedProcess {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [Parameter(Mandatory=$true)][string]$WorkingDirectory,
        [Parameter(Mandatory=$true)][string]$StdoutPath,
        [Parameter(Mandatory=$true)][string]$StderrPath,
        [Parameter(Mandatory=$true)][string]$ExitCodePath,
        [Parameter(Mandatory=$true)][string]$HeartbeatPath
    )

    if ($Simulate) {
        "SIMULATE FilePath=$FilePath Arguments=$(Join-CommandLine -ArgumentValues $Arguments)" | Set-Content -LiteralPath $StdoutPath -Encoding UTF8
        "0" | Set-Content -LiteralPath $ExitCodePath -Encoding UTF8
        ("simulated {0}" -f (Get-Date -Format "s")) | Set-Content -LiteralPath $HeartbeatPath -Encoding UTF8
        return 0
    }

    $runnerPath = [System.IO.Path]::ChangeExtension($HeartbeatPath, ".runner.cmd")
    $argLine = ($Arguments | ForEach-Object { Quote-CmdArg $_ }) -join " "
    $runnerLines = @(
        "@echo off",
        "cd /d $(Quote-CmdArg $WorkingDirectory)",
        "$(Quote-CmdArg $FilePath) $argLine 1> $(Quote-CmdArg $StdoutPath) 2> $(Quote-CmdArg $StderrPath)",
        "exit /b %ERRORLEVEL%"
    )
    Set-Content -LiteralPath $runnerPath -Value $runnerLines -Encoding ASCII

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "cmd.exe"
    $psi.Arguments = "/d /c $(Quote-CmdArg $runnerPath)"
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    try {
        [void]$proc.Start()

        while (-not $proc.WaitForExit(30000)) {
            ("running {0} pid={1}" -f (Get-Date -Format "s"), $proc.Id) | Set-Content -LiteralPath $HeartbeatPath -Encoding UTF8
            $proc.Refresh()
        }

        $code = $proc.ExitCode
        $code | Set-Content -LiteralPath $ExitCodePath -Encoding UTF8
        ("ended {0} exitcode={1}" -f (Get-Date -Format "s"), $code) | Set-Content -LiteralPath $HeartbeatPath -Encoding UTF8
        return $code
    }
    finally {
        $proc.Dispose()
    }
}

function Build-AsrCommand {
    param([object]$Job)
    $audioPath = Assert-AllowedPath -PathValue ([string]$Job.audio_path) -FieldName "audio_path"
    $outputDir = Assert-AllowedPath -PathValue ([string]$Job.output_dir) -FieldName "output_dir"
    $properNames = Assert-AllowedPath -PathValue ([string]$Job.proper_names) -FieldName "proper_names"
    $expectedCsv = Assert-AllowedPath -PathValue ([string]$Job.expected_csv) -FieldName "expected_csv"
    $script = Assert-AllowedPath -PathValue $AsrScript -FieldName "AsrScript"

    if (-not (Test-Path -LiteralPath $audioPath)) { throw "Audio introuvable: $audioPath" }
    if (-not (Test-Path -LiteralPath $script)) { throw "Script ASR introuvable: $script" }
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    $modelKey = [string]$Job.model_key
    if ([string]::IsNullOrWhiteSpace($modelKey)) {
        $modelKey = "Voxtral_Mini_3B_Transformers"
    }

    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $script,
        "-LocalAudioPath", $audioPath,
        "-ServerDropDir", $outputDir,
        "-LocalOutDir", $outputDir,
        "-ProperNamesFile", $properNames,
        "-ServerHostPort", $ServerHostPort,
        "-Subtitles",
        "-ModelKey", $modelKey
    )
    return @{
        FilePath = "powershell.exe"
        Arguments = $args
        WorkingDirectory = $outputDir
        ExpectedOutput = $expectedCsv
    }
}

function Build-AnnotationCommand {
    param([object]$Job)
    $infos = Assert-AllowedPath -PathValue ([string]$Job.infos_projet) -FieldName "infos_projet"
    $batch = Assert-AllowedPath -PathValue $AnnotationBatch -FieldName "AnnotationBatch"
    if (-not (Test-Path -LiteralPath $infos)) { throw "infos_projet introuvable: $infos" }
    if (-not (Test-Path -LiteralPath $batch)) { throw "Batch AnnotationPhotosGPT introuvable: $batch" }

    $allowedOptions = @(
        "--dry-run",
        "--limit",
        "--night",
        "--vlm-strict",
        "--reset-vlm",
        "--reset-llm",
        "--reset-vlm-plus",
        "--only-new-dictee",
        "--rerun-weak",
        "--rerun-weak-backend"
    )
    $options = @()
    if ($Job.options -is [System.Array]) {
        for ($i = 0; $i -lt $Job.options.Count; $i++) {
            $name = [string]$Job.options[$i]
            if ($allowedOptions -notcontains $name) { throw "Option batch refusee: $name" }
            if (($i + 1) -ge $Job.options.Count) { throw "Valeur manquante pour option batch: $name" }
            $value = [string]$Job.options[$i + 1]
            if ($value -match '[;&|<>"]') { throw "Valeur option refusee pour ${name}: $value" }
            if ($name -ne "--rerun-weak-backend" -and $value -notmatch '^\d+$') {
                throw "Valeur numerique attendue pour ${name}: $value"
            }
            if ($name -eq "--rerun-weak-backend" -and @("same", "local", "remote") -notcontains $value) {
                throw "Valeur invalide pour --rerun-weak-backend: $value"
            }
            $options += $name
            $options += $value
            $i++
        }
    }

    $args = @("/c", $batch, $infos) + $options
    return @{
        FilePath = "cmd.exe"
        Arguments = $args
        WorkingDirectory = (Split-Path $batch -Parent)
        ExpectedOutput = ""
    }
}

function Build-DeepSeekOcrSimulation {
    param([object]$Job)

    $projectId = Get-JobString -Job $Job -Name "project_id"
    if ([string]::IsNullOrWhiteSpace($projectId)) { throw "project_id manquant" }
    if ($projectId -notmatch '^\d{4}-J\d+$') { throw "project_id invalide: $projectId" }

    $sourcePdf = Assert-AllowedPath -PathValue (Get-JobString -Job $Job -Name "source_pdf") -FieldName "source_pdf"
    $outputDir = Assert-AllowedPath -PathValue (Get-JobString -Job $Job -Name "output_dir") -FieldName "output_dir"
    $pngDir = Assert-AllowedPath -PathValue (Get-JobString -Job $Job -Name "png_dir") -FieldName "png_dir"
    if ([IO.Path]::GetExtension($sourcePdf) -ne ".pdf") { throw "source_pdf doit etre un PDF: $sourcePdf" }

    $pages = Assert-Pages -Value (Get-JobValue -Job $Job -Name "pages")
    $tileCount = Assert-TileCount -Value (Get-JobValue -Job $Job -Name "tile_count")
    $postprocess = Assert-StrictBoolean -Value (Get-JobValue -Job $Job -Name "postprocess") -FieldName "postprocess"
    $retryGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $Job -Name "retry_glitch_pages") -FieldName "retry_glitch_pages"
    $tileGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $Job -Name "tile_glitch_pages") -FieldName "tile_glitch_pages"
    $fallbackTesseractPages = Assert-StrictBoolean -Value (Get-JobValue -Job $Job -Name "fallback_tesseract_pages") -FieldName "fallback_tesseract_pages"

    $leaf = [IO.Path]::GetFileNameWithoutExtension($sourcePdf)
    $expectedTxt = Join-Path $outputDir ("{0}.merged.txt" -f $leaf)
    $expectedMd = Join-Path $outputDir ("{0}.merged.md" -f $leaf)

    $payload = [ordered]@{
        input_dir = $pngDir
        output_dir = $outputDir
        pages = $pages
        tile_count = $tileCount
        postprocess = $postprocess
        retry_glitch_pages = $retryGlitchPages
        tile_glitch_pages = $tileGlitchPages
        fallback_tesseract_pages = $fallbackTesseractPages
    }

    return @{
        SimulatedJob = $true
        ProjectId = $projectId
        SourcePdf = $sourcePdf
        PngDir = $pngDir
        OutputDir = $outputDir
        Endpoint = $DeepSeekEndpoint
        Payload = $payload
        ExpectedTxt = $expectedTxt
        ExpectedMd = $expectedMd
    }
}

function Invoke-DeepSeekOcrSimulation {
    param(
        [Parameter(Mandatory=$true)][hashtable]$Simulation,
        [Parameter(Mandatory=$true)][string]$StdoutPath,
        [Parameter(Mandatory=$true)][string]$StderrPath,
        [Parameter(Mandatory=$true)][string]$ExitCodePath,
        [Parameter(Mandatory=$true)][string]$HeartbeatPath
    )

    $manifestPath = [System.IO.Path]::ChangeExtension($StdoutPath, ".manifest.json")
    $manifest = [ordered]@{
        simulated = $true
        project_id = $Simulation["ProjectId"]
        source_pdf = $Simulation["SourcePdf"]
        planned_png_dir = $Simulation["PngDir"]
        planned_output_dir = $Simulation["OutputDir"]
        planned_endpoint = $Simulation["Endpoint"]
        planned_payload = $Simulation["Payload"]
        expected_txt = $Simulation["ExpectedTxt"]
        expected_md = $Simulation["ExpectedMd"]
        created_at = (Get-Date -Format "s")
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    @(
        "SIMULATION deepseek_ocr uniquement",
        "Aucun appel HTTP vers $($Simulation["Endpoint"])",
        "Aucune conversion PDF vers PNG lancee",
        "Manifest: $manifestPath",
        "TXT final prevu: $($Simulation["ExpectedTxt"])",
        "MD final prevu: $($Simulation["ExpectedMd"])"
    ) | Set-Content -LiteralPath $StdoutPath -Encoding UTF8
    "" | Set-Content -LiteralPath $StderrPath -Encoding UTF8
    "0" | Set-Content -LiteralPath $ExitCodePath -Encoding UTF8
    ("simulated {0}" -f (Get-Date -Format "s")) | Set-Content -LiteralPath $HeartbeatPath -Encoding UTF8
    return 0
}

function Validate-Job {
    param([object]$Job)
    $jobId = Get-JobString -Job $Job -Name "job_id"
    $type = Get-JobString -Job $Job -Name "type"
    if ([string]::IsNullOrWhiteSpace($jobId)) { throw "job_id manquant" }
    if ([string]::IsNullOrWhiteSpace($type)) { throw "type manquant" }
    if ($type -eq "deepseek_ocr") {
        if ([string]::IsNullOrWhiteSpace((Get-JobString -Job $Job -Name "project_id"))) { throw "project_id manquant" }
        return
    }
    if ([string]::IsNullOrWhiteSpace((Get-JobString -Job $Job -Name "affaire"))) { throw "affaire manquante" }
    if ([string]::IsNullOrWhiteSpace((Get-JobString -Job $Job -Name "captation"))) { throw "captation manquante" }
}

function Recover-RunningJobs {
    param(
        [string]$RunningDir,
        [string]$QueuedDir,
        [string]$FailedDir,
        [string]$LogsDir,
        [string]$Policy
    )
    $runningJobs = Get-ChildItem -LiteralPath $RunningDir -Filter "*.json" -File -ErrorAction SilentlyContinue
    foreach ($item in $runningJobs) {
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeName = ([IO.Path]::GetFileNameWithoutExtension($item.Name)) -replace '[^A-Za-z0-9_.-]', '_'
        $logPath = Join-Path $LogsDir ("{0}_{1}.recovery.txt" -f $stamp, $safeName)
        if ($Policy -eq "ignore") {
            "running job left untouched: $($item.FullName)" | Set-Content -LiteralPath $logPath -Encoding UTF8
            continue
        }
        if ($Policy -eq "requeue") {
            $dest = Move-JobFile -Source $item.FullName -DestinationDir $QueuedDir
            "running job requeued after stale recovery: $dest" | Set-Content -LiteralPath $logPath -Encoding UTF8
            continue
        }
        $failed = Move-JobFile -Source $item.FullName -DestinationDir $FailedDir
        "running job marked failed after stale recovery: $failed" | Set-Content -LiteralPath $logPath -Encoding UTF8
    }
}

Ensure-JobDirs -Root $JobsRoot

$lockPath = Join-Path $JobsRoot "spooler.lock"
$lockStream = $null
try {
    $lockStream = Open-SpoolerLock -LockPath $lockPath -StaleMinutes $StaleLockMinutes
    if (-not $lockStream) { return }

    $lockText = [System.Text.Encoding]::UTF8.GetBytes(("pid={0} start={1}" -f $PID, (Get-Date -Format "s")))
    $lockStream.Write($lockText, 0, $lockText.Length)

    $queuedDir = Join-Path $JobsRoot "queued"
    $runningDir = Join-Path $JobsRoot "running"
    $doneDir = Join-Path $JobsRoot "done"
    $failedDir = Join-Path $JobsRoot "failed"
    $logsDir = Join-Path $JobsRoot "logs"

    Recover-RunningJobs -RunningDir $runningDir -QueuedDir $queuedDir -FailedDir $failedDir -LogsDir $logsDir -Policy $RecoverRunning

    $jobFile = Get-ChildItem -LiteralPath $queuedDir -Filter "*.json" -File |
        Sort-Object LastWriteTimeUtc |
        Select-Object -First 1

    if (-not $jobFile) {
        Write-Host "Aucun job en attente."
        return
    }

    $runningJobPath = Move-JobFile -Source $jobFile.FullName -DestinationDir $runningDir
    $job = Get-Content -LiteralPath $runningJobPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Validate-Job -Job $job

    $safeJobId = ([string]$job.job_id) -replace '[^A-Za-z0-9_.-]', '_'
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logBase = Join-Path $logsDir ("{0}_{1}" -f $stamp, $safeJobId)
    $stdoutPath = "$logBase.stdout.log"
    $stderrPath = "$logBase.stderr.log"
    $exitCodePath = "$logBase.exitcode.txt"
    $heartbeatPath = "$logBase.heartbeat.txt"
    $commandPath = "$logBase.command.txt"

    try {
        switch ([string]$job.type) {
            "asr_voxtral" {
                $cmd = Build-AsrCommand -Job $job
            }
            "annotation_photos_batch" {
                $cmd = Build-AnnotationCommand -Job $job
            }
            "deepseek_ocr" {
                $cmd = Build-DeepSeekOcrSimulation -Job $job
            }
            default {
                throw "Type de job refuse: $($job.type)"
            }
        }

        if ($cmd["SimulatedJob"]) {
            ($cmd | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $commandPath -Encoding UTF8
            $rc = Invoke-DeepSeekOcrSimulation `
                -Simulation $cmd `
                -StdoutPath $stdoutPath `
                -StderrPath $stderrPath `
                -ExitCodePath $exitCodePath `
                -HeartbeatPath $heartbeatPath
        }
        else {
            (@($cmd["FilePath"]) + $cmd["Arguments"]) -join " " | Set-Content -LiteralPath $commandPath -Encoding UTF8
            $rc = Invoke-LoggedProcess `
                -FilePath $cmd["FilePath"] `
                -Arguments $cmd["Arguments"] `
                -WorkingDirectory $cmd["WorkingDirectory"] `
                -StdoutPath $stdoutPath `
                -StderrPath $stderrPath `
                -ExitCodePath $exitCodePath `
                -HeartbeatPath $heartbeatPath
        }

        if ($rc -eq 0) {
            Move-JobFile -Source $runningJobPath -DestinationDir $doneDir | Out-Null
        }
        else {
            Move-JobFile -Source $runningJobPath -DestinationDir $failedDir | Out-Null
        }
    }
    catch {
        $_ | Out-String | Set-Content -LiteralPath $stderrPath -Encoding UTF8
        "999" | Set-Content -LiteralPath $exitCodePath -Encoding UTF8
        ("failed {0}" -f (Get-Date -Format "s")) | Set-Content -LiteralPath $heartbeatPath -Encoding UTF8
        if (Test-Path -LiteralPath $runningJobPath) {
            Move-JobFile -Source $runningJobPath -DestinationDir $failedDir | Out-Null
        }
        throw
    }
}
finally {
    if ($lockStream) {
        $lockStream.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}
