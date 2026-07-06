param(
    [Parameter(Mandatory=$true)][string]$JobPath,
    [string]$PythonPath = "D:\GPT4all_local\.venv\Scripts\python.exe",
    [string]$Endpoint = "http://127.0.0.1:5050/ocr_deepseek_batch",
    [int]$Dpi = 300,
    [int]$TimeoutSec = 900
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

function Invoke-JsonPost {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][hashtable]$Payload,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter(Mandatory=$true)][int]$Timeout
    )
    $body = $Payload | ConvertTo-Json -Depth 8
    try {
        $response = Invoke-WebRequest -Uri $Url -Method Post -Headers $Headers -ContentType "application/json; charset=utf-8" -Body $body -TimeoutSec $Timeout -UseBasicParsing
        return @{
            StatusCode = [int]$response.StatusCode
            Body = [string]$response.Content
            Error = ""
        }
    }
    catch {
        $status = 0
        $content = ""
        if ($_.Exception.Response) {
            try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $content = $reader.ReadToEnd()
                $reader.Dispose()
            }
            catch {
                $content = ""
            }
        }
        return @{
            StatusCode = $status
            Body = $content
            Error = ($_.Exception.Message)
        }
    }
}

function Get-LocalApiKey {
    $configPath = Assert-AllowedPath -PathValue "D:\GPT4all_local\config\config.json" -FieldName "config.json"
    if (-not (Test-Path -LiteralPath $configPath)) {
        throw "Configuration introuvable: $configPath"
    }
    $cfg = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $cfg.api_keys) {
        throw "api_keys absent de config.json"
    }
    foreach ($prop in $cfg.api_keys.PSObject.Properties) {
        $value = [string]$prop.Value
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }
    throw "Aucune cle API locale exploitable dans config.json"
}

$jobFile = Assert-AllowedPath -PathValue $JobPath -FieldName "JobPath"
$pythonExe = Assert-AllowedPath -PathValue $PythonPath -FieldName "PythonPath"
$converterScript = Assert-AllowedPath -PathValue (Join-Path $PSScriptRoot "run_deepseek_ocr_job.convert_png.ps1") -FieldName "converterScript"
if (-not (Test-Path -LiteralPath $pythonExe)) { throw "Python introuvable: $pythonExe" }
if (-not (Test-Path -LiteralPath $converterScript)) { throw "Helper conversion introuvable: $converterScript" }
if ($Endpoint -ne "http://127.0.0.1:5050/ocr_deepseek_batch") {
    throw "Endpoint refuse pour ce lot: $Endpoint"
}

$job = Get-Content -LiteralPath $jobFile -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-JobString -Job $job -Name "type") -ne "deepseek_ocr") {
    throw "Type de job refuse: $((Get-JobString -Job $job -Name "type"))"
}

$jobId = Get-JobString -Job $job -Name "job_id"
if ([string]::IsNullOrWhiteSpace($jobId)) { throw "job_id manquant" }

$sourcePdf = Assert-AllowedPath -PathValue (Get-JobString -Job $job -Name "source_pdf") -FieldName "source_pdf"
$pngDir = Assert-AllowedPath -PathValue (Get-JobString -Job $job -Name "png_dir") -FieldName "png_dir"
$outputDir = Assert-AllowedPath -PathValue (Get-JobString -Job $job -Name "output_dir") -FieldName "output_dir"
if ([IO.Path]::GetExtension($sourcePdf) -ne ".pdf") { throw "source_pdf doit etre un PDF: $sourcePdf" }
if (-not (Test-Path -LiteralPath $sourcePdf)) { throw "PDF source introuvable: $sourcePdf" }

$pages = Assert-Pages -Value (Get-JobValue -Job $job -Name "pages")
if ($pages.Count -ne 1) { throw "Lot 9B limite a une seule page" }
$tileCount = Assert-TileCount -Value (Get-JobValue -Job $job -Name "tile_count")
$postprocess = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "postprocess") -FieldName "postprocess"
$retryGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "retry_glitch_pages") -FieldName "retry_glitch_pages"
$tileGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "tile_glitch_pages") -FieldName "tile_glitch_pages"
$fallbackTesseractPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "fallback_tesseract_pages") -FieldName "fallback_tesseract_pages"
if ($tileCount -ne 2) { throw "Lot 9B: tile_count doit valoir 2" }
if (-not $postprocess -or -not $retryGlitchPages -or -not $tileGlitchPages -or $fallbackTesseractPages) {
    throw "Lot 9B: options OCR attendues postprocess=true retry=true tile=true fallback=false"
}

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $converterScript -JobPath $jobFile -PythonPath $pythonExe -Dpi $Dpi
if ($LASTEXITCODE -ne 0) {
    throw "Conversion PNG echouee avant appel batch, exitcode=$LASTEXITCODE"
}

$safeJobId = $jobId -replace '[^A-Za-z0-9_.-]', '_'
$pngManifestPath = Join-Path $pngDir ("{0}.png_manifest.json" -f $safeJobId)
if (-not (Test-Path -LiteralPath $pngManifestPath)) {
    throw "Manifest PNG introuvable: $pngManifestPath"
}
$pngManifest = Get-Content -LiteralPath $pngManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$imagePaths = @($pngManifest.image_paths)
if ($imagePaths.Count -ne 1) {
    throw "Lot 9B attend exactement une image PNG, obtenu: $($imagePaths.Count)"
}
foreach ($imagePath in $imagePaths) {
    $checked = Assert-AllowedPath -PathValue ([string]$imagePath) -FieldName "image_paths"
    if (-not (Test-Path -LiteralPath $checked)) {
        throw "PNG introuvable: $checked"
    }
}

try {
    $health = Invoke-WebRequest -Uri "http://127.0.0.1:5050/" -Method Get -TimeoutSec 5 -UseBasicParsing
    $serverStarted = $true
    $serverStatus = [int]$health.StatusCode
}
catch {
    if ($_.Exception.Response) {
        $serverStarted = $true
        try { $serverStatus = [int]$_.Exception.Response.StatusCode } catch { $serverStatus = 0 }
    }
    else {
        $serverStarted = $false
        $serverStatus = 0
    }
}
if (-not $serverStarted) {
    throw "Serveur Flask local indisponible sur http://127.0.0.1:5050/"
}

$payload = @{
    image_paths = @($imagePaths)
    output_dir = $outputDir
    postprocess = $true
    retry_glitch_pages = $true
    tile_glitch_pages = $true
    fallback_tesseract_pages = $false
    tile_count = 2
}

$apiKey = Get-LocalApiKey
$headers = @{ "x-api-key" = $apiKey }
$response = Invoke-JsonPost -Url $Endpoint -Payload $payload -Headers $headers -Timeout $TimeoutSec
$responsePath = Join-Path $outputDir ("{0}.ocr_deepseek_batch.response.json" -f $safeJobId)
$diagnosticsPath = Join-Path $outputDir ("{0}.ocr_deepseek_batch.diagnostics.json" -f $safeJobId)

$response.Body | Set-Content -LiteralPath $responsePath -Encoding UTF8
$parsed = $null
if (-not [string]::IsNullOrWhiteSpace($response.Body)) {
    try { $parsed = $response.Body | ConvertFrom-Json } catch { $parsed = $null }
}

$textPaths = @()
if ($parsed -and $parsed.results) {
    foreach ($result in @($parsed.results)) {
        foreach ($field in @("text_path", "raw_text_path", "clean_text_path", "attempt1_text_path", "attempt2_text_path", "tiled_clean_text_path", "fallback_tesseract_text_path")) {
            $prop = $result.PSObject.Properties[$field]
            if ($prop -and $prop.Value) {
                $textPaths += [string]$prop.Value
            }
        }
    }
}
$textPaths = @($textPaths | Select-Object -Unique)

$diagnostics = [ordered]@{
    job_id = $jobId
    endpoint = $Endpoint
    server_preflight_ok = $serverStarted
    server_preflight_status = $serverStatus
    source_pdf = $sourcePdf
    png_manifest = $pngManifestPath
    payload = $payload
    auth_header_sent = $true
    http_status = $response.StatusCode
    http_error = $response.Error
    response_json_path = $responsePath
    response_parse_ok = ($null -ne $parsed)
    response_ok = if ($parsed -and $parsed.PSObject.Properties["ok"]) { [bool]$parsed.ok } else { $false }
    text_paths = $textPaths
    text_paths_existing = @($textPaths | Where-Object { Test-Path -LiteralPath $_ })
    timestamp = (Get-Date -Format "s")
}
$diagnostics | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $diagnosticsPath -Encoding UTF8

Write-Host "DeepSeekOCR batch call completed."
Write-Host "Response: $responsePath"
Write-Host "Diagnostics: $diagnosticsPath"

if ($response.StatusCode -lt 200 -or $response.StatusCode -ge 300) {
    throw "Appel /ocr_deepseek_batch echoue HTTP $($response.StatusCode): $($response.Error)"
}
if (-not $parsed) {
    throw "Reponse /ocr_deepseek_batch non JSON"
}
if (-not [bool]$parsed.ok) {
    throw "Reponse /ocr_deepseek_batch ok=false"
}
