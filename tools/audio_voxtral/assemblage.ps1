#requires -version 5.1

param(
  [string]$WorkDir = (Get-Location).Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$WorkDir = $WorkDir.Trim('"')
Set-Location -LiteralPath $WorkDir

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

$ListFile = "liste.txt"
$LogFile  = "assemblage_log.txt"

function Log($msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] $msg"
    $line | Tee-Object -FilePath $LogFile -Append
}

if (-not (Get-Command ffmpeg -ErrorAction SilentlyContinue)) {
    throw "ffmpeg est introuvable dans le PATH."
}
if (-not (Test-Path -LiteralPath $ListFile)) {
    throw "Fichier introuvable: $ListFile"
}

$lines = Get-Content -LiteralPath $ListFile -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }
if ($lines.Count -lt 2) {
    throw "liste.txt doit contenir au moins 2 lignes au format: file '...'."
}

if ($lines[0] -notmatch "^\s*file\s+'(.+)'\s*$") {
    throw "Première ligne invalide: $($lines[0])"
}
$SrcFile = $Matches[1]
if (-not (Test-Path -LiteralPath $SrcFile)) {
    throw "Le fichier source (1ère ligne) est introuvable: $SrcFile"
}

$missing = @()
foreach ($l in $lines) {
    if ($l -notmatch "^\s*file\s+'(.+)'\s*$") {
        throw "Ligne invalide dans liste.txt: $l"
    }
    $p = $Matches[1]
    if (-not (Test-Path -LiteralPath $p)) { $missing += $p }
}
if ($missing.Count -gt 0) {
    throw "Fichiers manquants:`n- " + ($missing -join "`n- ")
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($SrcFile)
$baseName = $baseName -replace "\s+partie\s+\d+.*$", ""
$baseName = $baseName -replace "\s+", "_"
$OutFile  = "${baseName}_complet.wav"

if (Test-Path -LiteralPath $OutFile) {
    throw "Le fichier de sortie existe déjà (refus d'écrasement): $OutFile"
}

Remove-Item -LiteralPath $LogFile -ErrorAction SilentlyContinue
Log "Début assemblage"
Log "Liste: $ListFile"
Log "Source horodatages (1ère ligne): $SrcFile"
Log "Sortie: $OutFile"

Log "Concaténation ffmpeg (sans ré-encodage)"
& ffmpeg -hide_banner -loglevel error -f concat -safe 0 -i $ListFile -c copy $OutFile
if ($LASTEXITCODE -ne 0) { throw "ffmpeg a échoué (code=$LASTEXITCODE)." }
if (-not (Test-Path -LiteralPath $OutFile)) { throw "Sortie non créée: $OutFile" }
Log "Concaténation terminée"

$src = Get-Item -LiteralPath $SrcFile
$dst = Get-Item -LiteralPath $OutFile

Log ("Avant horodatages sortie: Creation={0} Write={1} Access={2}" -f $dst.CreationTime, $dst.LastWriteTime, $dst.LastAccessTime)

$dst.CreationTime    = $src.CreationTime
$dst.LastWriteTime   = $src.LastWriteTime
$dst.LastAccessTime  = $src.LastAccessTime

$dst = Get-Item -LiteralPath $OutFile
Log ("Après horodatages sortie: Creation={0} Write={1} Access={2}" -f $dst.CreationTime, $dst.LastWriteTime, $dst.LastAccessTime)

$size = (Get-Item -LiteralPath $OutFile).Length
Log "Taille sortie (octets): $size"

$hash = Get-FileHash -LiteralPath $OutFile -Algorithm SHA256
Log "SHA-256 sortie: $($hash.Hash)"

Log "Fin assemblage OK"

Write-Host ""
Write-Host "OK -> $OutFile"
Write-Host "Log -> $LogFile"
