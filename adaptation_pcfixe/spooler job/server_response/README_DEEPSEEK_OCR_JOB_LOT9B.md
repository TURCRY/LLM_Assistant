# Lot 9B - DeepSeekOCR spooler, appel batch reel borne

Date : 2026-07-07

## Portee

Ce lot ajoute une version draft qui :

1. convertit le PDF guide en PNG comme au lot 9A ;
2. appelle reellement Flask local :

```text
http://127.0.0.1:5050/ocr_deepseek_batch
```

3. ecrit la reponse JSON et un diagnostic JSON ;
4. laisse la fusion finale `.md` / `.txt` complete pour un lot suivant.

Ce qui n'a pas ete modifie :

- `app.py` ;
- Flask ;
- la tache planifiee ;
- les scripts de production dans `D:\GPT4all_local\scripts\jobs`.

## Fichiers produits dans server_response

```text
D:\GPT4all_local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.call_batch.ps1
D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_call_batch.ps1
D:\GPT4all_local\proposed_changes\spooler job\server_response\README_DEEPSEEK_OCR_JOB_LOT9B.md
```

## Contrat Flask utilise

La route `/ocr_deepseek_batch` attend :

```json
{
  "image_paths": ["...\\page_0001.png"],
  "output_dir": "...",
  "postprocess": true,
  "retry_glitch_pages": true,
  "tile_glitch_pages": true,
  "fallback_tesseract_pages": false,
  "tile_count": 2
}
```

La route exige aussi `x-api-key`. Le helper lit la cle dans :

```text
D:\GPT4all_local\config\config.json
```

La cle n'est pas ecrite dans les logs ni dans les diagnostics.

## Fonctionnement

`run_pending_jobs.deepseek_call_batch.ps1` conserve :

- le verrou `spooler.lock` ;
- un seul job a la fois ;
- la reprise `running -> failed` ;
- les branches existantes ASR et annotation ;
- les logs stdout / stderr / exitcode / heartbeat.

Pour `deepseek_ocr`, il appelle :

```text
run_deepseek_ocr_job.call_batch.ps1
```

Le helper :

- valide les chemins sous `C:\Affaires` et `D:\GPT4all_local` ;
- refuse tout endpoint autre que `http://127.0.0.1:5050/ocr_deepseek_batch` ;
- limite ce lot a une seule page ;
- impose `tile_count = 2` ;
- impose `postprocess=true`, `retry_glitch_pages=true`, `tile_glitch_pages=true`, `fallback_tesseract_pages=false` ;
- convertit le PDF en PNG via le helper 9A ;
- verifie que Flask local repond ;
- appelle `/ocr_deepseek_batch` ;
- ecrit la reponse JSON et les diagnostics dans `output_dir`.

## Tests

### 1. Validation syntaxique PowerShell

```text
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.call_batch.ps1
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_call_batch.ps1
```

### 2. Job depose

Job cree en queue :

```text
C:\Affaires\_jobs\queued\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.json
```

Parametres utiles :

- `project_id = 2025-J47`
- `pages = [8]`
- `tile_count = 2`
- `postprocess = true`
- `retry_glitch_pages = true`
- `tile_glitch_pages = true`
- `fallback_tesseract_pages = false`

### 3. Commande lancee

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_call_batch.ps1" -JobsRoot "C:\Affaires\_jobs"
```

### 4. Resultat spooler

Resultat :

- `queued` : job absent apres execution ;
- `running` : job absent apres execution ;
- `done` : job present ;
- `failed` : job absent ;
- exit code : `0` ;
- stderr spooler : vide.

Job final :

```text
C:\Affaires\_jobs\done\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.json
```

Logs spooler :

```text
C:\Affaires\_jobs\logs\20260707_002101_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.command.txt
C:\Affaires\_jobs\logs\20260707_002101_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.stdout.log
C:\Affaires\_jobs\logs\20260707_002101_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.stderr.log
C:\Affaires\_jobs\logs\20260707_002101_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.exitcode.txt
C:\Affaires\_jobs\logs\20260707_002101_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.heartbeat.txt
```

## Sorties produites

PNG manifest :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\102 Assignation délivrée à 3F_pages_user_8\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.png_manifest.json
```

PNG utilise :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\102 Assignation délivrée à 3F_pages_user_8\page_0001.png
```

Reponse Flask :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.ocr_deepseek_batch.response.json
```

Diagnostics :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_lot9b_20260707_001500.ocr_deepseek_batch.diagnostics.json
```

Fichiers texte DeepSeek produits :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8\0001_page_0001.deepseek_ocr.txt
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8\0001_page_0001.deepseek_ocr.raw.txt
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8\0001_page_0001.deepseek_ocr.clean.txt
```

Verification integrite basique :

```text
0001_page_0001.deepseek_ocr.txt       562 caracteres lisibles
0001_page_0001.deepseek_ocr.raw.txt   562 caracteres lisibles
0001_page_0001.deepseek_ocr.clean.txt 562 caracteres lisibles
```

## Reponse et diagnostics

Diagnostic essentiel :

```text
server_preflight_ok = true
server_preflight_status = 404
http_status = 200
response_parse_ok = true
response_ok = true
auth_header_sent = true
```

Note : le preflight utilise `GET /`, qui retourne 404 sur ce serveur. Cela prouve tout de meme que Flask local repond. L'appel utile `/ocr_deepseek_batch` a bien retourne `HTTP 200`.

La reponse Flask indique :

```text
ok = true
engine = deepseek_ocr
pages = 1
metrics.count = 1
metrics.total_text_chars = 554
retry_used_count = 0
tile_used_count = 0
fallback_tesseract_used_count = 0
```

## Limites

- Lot limite volontairement a une seule page J47.
- Pas de fusion finale `.md` / `.txt`.
- Pas de relance automatique en cas d'echec.
- Les PNG existants ne sont pas nettoyes avant conversion.
- Le preflight serveur pourrait etre ameliore avec une route de sante explicite si elle est ajoutee plus tard.

## Prochaine etape recommandee

Lot 9C :

- ajouter une fusion finale minimale `.txt` et `.md` ;
- ecrire dans le manifest final les chemins de sorties fusionnees ;
- confirmer la convention de nommage finale attendue par le laptop ;
- seulement apres validation, preparer une proposition de copie vers `D:\GPT4all_local\scripts\jobs`.
