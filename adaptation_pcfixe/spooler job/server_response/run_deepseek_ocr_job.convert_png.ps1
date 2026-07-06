param(
    [Parameter(Mandatory=$true)][string]$JobPath,
    [string]$PythonPath = "D:\GPT4all_local\.venv\Scripts\python.exe",
    [int]$Dpi = 300
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

$jobFile = Assert-AllowedPath -PathValue $JobPath -FieldName "JobPath"
$pythonExe = Assert-AllowedPath -PathValue $PythonPath -FieldName "PythonPath"
if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python introuvable: $pythonExe"
}
if ($Dpi -lt 72 -or $Dpi -gt 600) {
    throw "DPI refuse: $Dpi"
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
if ([IO.Path]::GetExtension($sourcePdf) -ne ".pdf") {
    throw "source_pdf doit etre un PDF: $sourcePdf"
}
if (-not (Test-Path -LiteralPath $sourcePdf)) {
    throw "PDF source introuvable: $sourcePdf"
}

$pages = Assert-Pages -Value (Get-JobValue -Job $job -Name "pages")
[void](Assert-TileCount -Value (Get-JobValue -Job $job -Name "tile_count"))
[void](Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "postprocess") -FieldName "postprocess")
[void](Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "retry_glitch_pages") -FieldName "retry_glitch_pages")
[void](Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "tile_glitch_pages") -FieldName "tile_glitch_pages")
[void](Assert-StrictBoolean -Value (Get-JobValue -Job $job -Name "fallback_tesseract_pages") -FieldName "fallback_tesseract_pages")

New-Item -ItemType Directory -Path $pngDir -Force | Out-Null

$safeJobId = $jobId -replace '[^A-Za-z0-9_.-]', '_'
$manifestPath = Join-Path $pngDir ("{0}.png_manifest.json" -f $safeJobId)
$pagesJson = ConvertTo-Json -InputObject @($pages) -Compress
$converterPath = Join-Path ([IO.Path]::GetTempPath()) ("deepseek_ocr_convert_png_{0}.py" -f ([guid]::NewGuid().ToString("N")))

$pythonCode = @'
import argparse
import json
from datetime import datetime
from pathlib import Path

import fitz


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--source-pdf", required=True)
    parser.add_argument("--png-dir", required=True)
    parser.add_argument("--pages-json", required=True)
    parser.add_argument("--dpi", required=True, type=int)
    parser.add_argument("--manifest-path", required=True)
    args = parser.parse_args()

    source_pdf = Path(args.source_pdf)
    png_dir = Path(args.png_dir)
    pages = json.loads(args.pages_json)
    png_dir.mkdir(parents=True, exist_ok=True)

    image_paths = []
    with fitz.open(str(source_pdf)) as doc:
        page_count = doc.page_count
        for index, page_number in enumerate(pages, start=1):
            if 1 <= page_number <= page_count:
                doc_index = page_number - 1
            elif page_count == len(pages):
                doc_index = index - 1
            else:
                raise ValueError(f"page hors limites: {page_number} / {page_count}")
            page = doc.load_page(doc_index)
            pix = page.get_pixmap(dpi=args.dpi, alpha=False)
            out_path = png_dir / f"page_{index:04d}.png"
            pix.save(str(out_path))
            image_paths.append(str(out_path))

    manifest = {
        "job_id": args.job_id,
        "source_pdf": str(source_pdf),
        "pages": pages,
        "png_dir": str(png_dir),
        "image_paths": image_paths,
        "dpi": args.dpi,
        "created_count": len(image_paths),
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }
    Path(args.manifest_path).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
'@

try {
    Set-Content -LiteralPath $converterPath -Value $pythonCode -Encoding UTF8
    $args = @(
        $converterPath,
        "--job-id", $jobId,
        "--source-pdf", $sourcePdf,
        "--png-dir", $pngDir,
        "--pages-json", $pagesJson,
        "--dpi", ([string]$Dpi),
        "--manifest-path", $manifestPath
    )
    & $pythonExe @args
    if ($LASTEXITCODE -ne 0) {
        throw "Conversion PDF vers PNG echouee, exitcode=$LASTEXITCODE"
    }
}
finally {
    Remove-Item -LiteralPath $converterPath -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Manifest PNG introuvable apres conversion: $manifestPath"
}

Write-Host "Conversion PDF vers PNG terminee."
Write-Host "Manifest: $manifestPath"
Write-Host "OutputDir OCR reserve non utilise dans ce lot: $outputDir"
