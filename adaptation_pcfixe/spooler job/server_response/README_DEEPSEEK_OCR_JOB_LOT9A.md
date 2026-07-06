# Lot 9A - DeepSeekOCR spooler, conversion PDF vers PNG

Date : 2026-07-07

## Portee

Ce lot ajoute une implementation reelle partielle pour le job `deepseek_ocr` :

```text
PDF source -> PNG page par page
```

Ce qui reste volontairement exclu :

- aucun appel a `http://127.0.0.1:5050/ocr_deepseek_batch` ;
- aucun lancement DeepSeekOCR reel ;
- aucune modification Flask ;
- aucune modification `app.py` ;
- aucune modification de la tache planifiee ;
- aucune modification des scripts de production dans `D:\GPT4all_local\scripts\jobs`.

## Fichiers produits dans server_response

```text
D:\GPT4all_local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.convert_png.ps1
D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_convert_png.ps1
D:\GPT4all_local\proposed_changes\spooler job\server_response\README_DEEPSEEK_OCR_JOB_LOT9A.md
```

## Fonctionnement propose

`run_pending_jobs.deepseek_convert_png.ps1` conserve le fonctionnement du spooler :

- un seul job a la fois ;
- verrou `spooler.lock` ;
- reprise `running -> failed` conservee ;
- dossiers `queued`, `running`, `done`, `failed`, `logs` ;
- branches existantes `asr_voxtral` et `annotation_photos_batch` conservees.

Pour `deepseek_ocr`, le spooler appelle :

```text
run_deepseek_ocr_job.convert_png.ps1
```

Le helper :

- valide les chemins sous `C:\Affaires` ou `D:\GPT4all_local` ;
- refuse les commandes arbitraires ;
- valide `pages`, `tile_count` et les booleens OCR, meme si les options OCR ne sont pas encore utilisees ;
- verifie l'existence du PDF source ;
- cree `png_dir` ;
- appelle `D:\GPT4all_local\.venv\Scripts\python.exe` ;
- utilise PyMuPDF / `fitz` ;
- convertit uniquement les pages demandees ;
- nomme les images `page_0001.png`, `page_0002.png`, etc. ;
- utilise `dpi = 300` ;
- ecrit un manifest JSON dans `png_dir`.

Manifest attendu :

```json
{
  "job_id": "...",
  "source_pdf": "...",
  "pages": [8],
  "png_dir": "...",
  "image_paths": ["...\\page_0001.png"],
  "dpi": 300,
  "created_count": 1,
  "timestamp": "..."
}
```

## Point traite : PDF guide mono-page

Le job laptop J47 garde `pages = [8]`, mais le PDF source guide :

```text
102 Assignation délivrée à 3F_pages_user_8.pdf
```

ne contient physiquement qu'une page.

Le helper gere donc ce cas :

- si le numero demande existe dans le PDF, il convertit cette page ;
- sinon, si le PDF contient exactement autant de pages que la liste `pages`, il convertit par ordre d'extraction ;
- le manifest conserve toujours les pages utilisateur originales, par exemple `[8]`.

Cela permet de garder le sens metier du job tout en traitant correctement un PDF guide deja extrait.

## Tests realises

### 1. Validation syntaxique PowerShell

```text
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.convert_png.ps1
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_convert_png.ps1
```

### 2. Preflight environnement

Python venv :

```text
D:\GPT4All_Local\.venv\Scripts\python.exe
```

PyMuPDF disponible :

```text
PyMuPDF 1.26.4
```

PDF source present :

```text
C:\Affaires\2025-J47\AA_Expert_Admin\Depot_initial\_Guided_Analysis\102 Assignation délivrée à 3F_pages_user_8.pdf
```

### 3. Premier test et correction

Un premier job Lot 9A a ete lance :

```text
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_20260707_000000
```

Resultat :

```text
failed
```

Cause :

```text
ValueError: page hors limites: 8 / 1
```

Analyse : le PDF guide contient une page physique, tandis que le job conserve le numero utilisateur `[8]`.

Correction appliquee dans le helper draft : prise en charge des PDF guides deja extraits, avec conversion par ordre lorsque `page_count == len(pages)`.

### 4. Test retry reussi

Job de retry cree en queue :

```text
C:\Affaires\_jobs\queued\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.json
```

Commande lancee :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_convert_png.ps1" -JobsRoot "C:\Affaires\_jobs"
```

Resultat :

- `queued` : job absent apres execution ;
- `running` : job absent apres execution ;
- `done` : job present ;
- `failed` : job absent ;
- exit code : `0` ;
- `stderr` : vide.

Job final :

```text
C:\Affaires\_jobs\done\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.json
```

Logs crees :

```text
C:\Affaires\_jobs\logs\20260707_000731_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.command.txt
C:\Affaires\_jobs\logs\20260707_000731_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.stdout.log
C:\Affaires\_jobs\logs\20260707_000731_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.stderr.log
C:\Affaires\_jobs\logs\20260707_000731_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.exitcode.txt
C:\Affaires\_jobs\logs\20260707_000731_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.heartbeat.txt
```

PNG cree :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\102 Assignation délivrée à 3F_pages_user_8\page_0001.png
```

Taille constatee :

```text
186239 octets
```

Manifest cree :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\102 Assignation délivrée à 3F_pages_user_8\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000.png_manifest.json
```

Contenu essentiel du manifest :

```json
{
  "job_id": "deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9a_retry_20260707_001000",
  "pages": [8],
  "image_paths": [
    "C:\\Affaires\\2025-J47\\AD_Expert_Traitements\\_OCR_Dire_Bordereau\\_pages_png\\102 Assignation délivrée à 3F_pages_user_8\\page_0001.png"
  ],
  "dpi": 300,
  "created_count": 1
}
```

## Verifications negatives

Aucun appel a `/ocr_deepseek_batch` n'est present dans la commande executee.

Le script 9A ne contient pas d'appel HTTP.

Aucun fichier OCR final `.txt` ou `.md` n'a ete cree. Le dossier de sortie OCR attendu reste absent :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8
```

## Limites

- Le premier job Lot 9A est reste dans `failed` parce qu'il a servi a detecter le cas PDF guide mono-page.
- Le helper cree les PNG mais ne nettoie pas les PNG existants si le meme job est relance.
- Les options OCR `postprocess`, `retry_glitch_pages`, `tile_glitch_pages`, `fallback_tesseract_pages` sont validees mais non utilisees dans ce lot.
- Le contrat exact du futur payload Flask reste a figer au lot suivant.

## Prochaine etape recommandee

Lot 9B :

- appeler reellement `http://127.0.0.1:5050/ocr_deepseek_batch` depuis un helper borne ;
- utiliser les `image_paths` du manifest PNG ;
- conserver l'absence de commande libre dans le JSON ;
- ecrire les sorties OCR `.txt` / `.md` ;
- fusionner les resultats finaux ;
- deplacer le job vers `done` ou `failed` selon le code retour.
