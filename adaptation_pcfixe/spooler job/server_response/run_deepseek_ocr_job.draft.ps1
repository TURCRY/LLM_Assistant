param(
    [Parameter(Mandatory=$true)][string]$JobPath,
    [string]$DeepSeekEndpoint = "http://127.0.0.1:5050/ocr_deepseek_batch"
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
        throw "pages doit etre une liste non vide"
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

$jobFile = Assert-AllowedPath -PathValue $JobPath -FieldName "JobPath"
$job = Get-Content -LiteralPath $jobFile -Raw -Encoding UTF8 | ConvertFrom-Json

if ([string]$job.type -ne "deepseek_ocr") {
    throw "Type de job refuse: $($job.type)"
}

$sourcePdf = Assert-AllowedPath -PathValue ([string]$job.source_pdf) -FieldName "source_pdf"
$outputDir = Assert-AllowedPath -PathValue ([string]$job.output_dir) -FieldName "output_dir"
$pngDir = Assert-AllowedPath -PathValue ([string]$job.png_dir) -FieldName "png_dir"
if ([IO.Path]::GetExtension($sourcePdf) -ne ".pdf") {
    throw "source_pdf doit etre un PDF: $sourcePdf"
}

$pages = Assert-Pages -Value (Get-JobValue -Job $job -Name "pages")
$tileCount = Assert-TileCount -Value (Get-JobValue -Job $job -Name "tile_count")
$postprocess = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "postprocess") -FieldName "postprocess"
$retryGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "retry_glitch_pages") -FieldName "retry_glitch_pages"
$tileGlitchPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "tile_glitch_pages") -FieldName "tile_glitch_pages"
$fallbackTesseractPages = Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "fallback_tesseract_pages") -FieldName "fallback_tesseract_pages"

$leaf = [IO.Path]::GetFileNameWithoutExtension($sourcePdf)
$plan = [ordered]@{
    simulated = $true
    conversion_prevue = "PDF vers PNG, une image par page demandee"
    source_pdf = $sourcePdf
    pages = $pages
    png_dir = $pngDir
    png_attendus = @($pages | ForEach-Object { Join-Path $pngDir ("page_{0:D4}.png" -f $_) })
    endpoint_prevu = $DeepSeekEndpoint
    payload_prevu = [ordered]@{
        input_dir = $pngDir
        output_dir = $outputDir
        pages = $pages
        tile_count = $tileCount
        postprocess = $postprocess
        retry_glitch_pages = $retryGlitchPages
        tile_glitch_pages = $tileGlitchPages
        fallback_tesseract_pages = $fallbackTesseractPages
    }
    fichiers_finaux_attendus = [ordered]@{
        txt = Join-Path $outputDir ("{0}.merged.txt" -f $leaf)
        md = Join-Path $outputDir ("{0}.merged.md" -f $leaf)
    }
}

$planPath = Join-Path (Split-Path $jobFile -Parent) ("{0}.deepseek_ocr_helper.simulation.json" -f $leaf)
$plan | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $planPath -Encoding UTF8
Write-Host "Simulation uniquement. Plan ecrit: $planPath"
