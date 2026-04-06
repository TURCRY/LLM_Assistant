# Génération de photos réduites & tableaux depuis `JPG\`

## Objet

Automatiser, **depuis un dossier `JPG\`**, la création :

1. d’un sous-dossier **`JPG reduit\`** contenant les JPEG compressés (EXIF/ICC conservés) ;
2. d’un **CSV** `photos.csv` (séparateur `;`, UTF-8-SIG) et d’un **XLS** `photos.xls` dans le **dossier parent**.

---

## Prérequis

* **Python 3.9+** disponible (`python --version`).
* Dépendances :

  ```powershell
  pip install pillow piexif
  ```
* (Optionnel pour `.xls` formaté) **Excel + pywin32** :

  ```powershell
  pip install pywin32
  ```

---

## Lancement

### Option A — Terminal (recommandé)

Se placer **dans** le dossier `JPG\`, puis :

```powershell
python gen_tout_depuis_JPG.py
```

### Option B — Double-clic (lanceur `.bat`)

Créez dans `JPG\` un fichier **`run_gen_tout_depuis_JPG.bat`** avec le contenu ci‑dessous, puis double‑cliquez‑le.

```bat
@echo off
setlocal EnableExtensions EnableDelayedExpansion
for %%I in ("%CD%") do set "CURR=%%~nI"
if /I not "%CURR%"=="JPG" (
  echo [ERREUR] Ce lanceur doit etre execute DEPUIS le dossier JPG^.
  echo Dossier courant: %CD%
  pause & exit /b 1
)
set "HERE=%~dp0"
set "SCRIPT=%HERE%gen_tout_depuis_JPG.py"
if not exist "%SCRIPT%" set "SCRIPT=C:\DevTools\Compression photo\gen_tout_depuis_JPG.py"
if not exist "%SCRIPT%" (
  echo [ERREUR] Script introuvable. & pause & exit /b 1
)
set "PY=" & where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
  echo [ERREUR] Python 3 introuvable dans le PATH. & pause & exit /b 1
)
for %%I in ("%CD%") do set "PARENT=%%~dpI"
set "ts=%date%_%time%" & set "ts=%ts:/=-%" & set "ts=%ts::=-%" & set "ts=%ts: =_%" & set "ts=%ts:.=-%"
set "LOGFILE=%PARENT%gen_tout_depuis_JPG_%ts%.log"
echo [INFO] Execution : %PY% "%SCRIPT%"
call %PY% "%SCRIPT%" 1>>"%LOGFILE%" 2>&1
if "%ERRORLEVEL%"=="0" (
  echo [OK] Termine (voir "%LOGFILE%").
  if exist "..\photos.csv" start "" "..\photos.csv"
  if exist "..\photos.xls" start "" "..\photos.xls"
) else (
  echo [ECHEC] Code retour: %ERRORLEVEL% (voir "%LOGFILE%") & pause
)
endlocal
```

---

## Paramètres (dans le script)

* `QUALITY = 63`  (ajustable ; 62–64 ≈ votre lot ; 80–85 = très bon visuel)
* `SUBSAMPLING = 2`  (`2` = 4:2:0 ; `1` = 4:2:2 ; `0` = 4:4:4)
* `PROGRESSIVE = True`  (JPEG progressif)
* Dossiers de sortie : `..\JPG reduit\`, `..\photos.csv`, `..\photos.xls`.

---

## Format des sorties

### 1) Réductions JPEG

* Images écrites dans **`..\JPG reduit\`** avec le **même nom** que l’original.
* **EXIF** (dont `DateTimeOriginal`, `Orientation`) et **profil ICC** **conservés**.

### 2) `photos.csv`

* Emplacement : **dossier parent** du `JPG\`.
* **Séparateur** : `;` (compat. Excel FR).
* **Encodage** : UTF‑8‑SIG (facilite l’ouverture directe dans Excel).
* **Colonnes (source/format)** :

  1. `nom_fichier_image` : nom du fichier (ex. `P1050638.JPG`).
  2. `horodatage_photo` : **EXIF** `DateTimeOriginal` prioritaire ; sinon `DateTimeDigitized/CreateDate` ; sinon `DateTime`. **Formaté** `jj/mm/aaaa hh:mm:ss`. *Pas de repli sur la date système*.
  3. `orientation_photo` : **EXIF Orientation (274)** converti **en degrés** : 1→0, 6→90, 8→270 (robustesse : 3→180 ; 2/4/5/7 mappés à 0/180/90/270 en ignorant le miroir).
  4. `chemin_photo_native` : chemin absolu de `JPG\` **avec `\` final**.
  5. `chemin_photo_reduite` : chemin absolu de `JPG reduit\` **avec `\` final**.
     6–10. `horodatage_secondes`, `t_audio`, `decalage_individuel`, `synchro_audio`, `decalage_moyen` : **vides** (réservés).

### 3) `photos.xls`

* Généré via Excel COM (si disponible) :

  * **format personnalisé** appliqué sur `horodatage_photo` : `jj/mm/aaaa hh:mm:ss` ;
  * séparateurs numériques FR (décimale `,`, milliers espace).
* **Fallback** (si Excel absent) : copie du CSV sous extension `.xls` (ouvrable, sans style).

---

## Dépannage

* **“Lancez ce script depuis le dossier 'JPG'”** : ouvrez le terminal **dans** `JPG\`.
* **`.xls` non formaté** : installer Excel + `pywin32`. Sinon le fallback crée un `.xls` simple.
* **Accents / séparateur mal interprétés** : CSV en `;` + UTF‑8‑SIG → ouvrir via double‑clic ou `Données > À partir d’un fichier texte/CSV` en précisant `;`.
* **Images non modifiées** : les originaux ne sont **jamais** modifiés ; tout est écrit dans `JPG reduit\`.
* **Chemins avec espaces** : toujours entourer de guillemets (`"..."`) dans les commandes.

---

## Contrôles complémentaires (ExifTool, optionnel)

Pour vérifier sous‑échantillonnage et estimation de qualité sur les **réduits** :

```powershell
& "C:\DevTools\Exiftools\exiftool-13.33_64\exiftool.exe" -csv -n -ext jpg -ext jpeg ^
  -FileName -JPEGQualityEstimate -YCbCrSubSampling "..\JPG reduit" ^
  | Out-File -FilePath "..\analyse_jpeg_reduits_params.csv" -Encoding utf8
```

* `YCbCrSubSampling` : `2 2` = **4:2:0** ; `2 1` = 4:2:2 ; `1 1` = 4:4:4.
* `JPEGQualityEstimate` peut être `<unknown>` selon l’encodeur (comportement normal).
