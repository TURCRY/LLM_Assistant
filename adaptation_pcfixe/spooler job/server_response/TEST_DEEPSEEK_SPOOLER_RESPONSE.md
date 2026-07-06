# Test spooler DeepSeekOCR - reponse PC fixe

Date du test : 2026-07-06

## Portee

Test demande : verifier que le spooler draft peut traiter en simulation un job `deepseek_ocr` cree par le laptop.

Contraintes respectees :

- aucun fichier de production dans `D:\GPT4all_local\scripts\jobs` n'a ete modifie ;
- Flask n'a pas ete modifie ;
- DeepSeekOCR reel n'a pas ete lance ;
- aucun appel HTTP vers `/ocr_deepseek_batch` n'a ete effectue ;
- seul le spooler draft `run_pending_jobs.deepseek_draft.ps1` a ete lance avec `-Simulate`.

## Fichiers lus

```text
D:\GPT4All_Local\proposed_changes\spooler job\server_response\README_DEEPSEEK_OCR_JOB_RESPONSE.md
D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_draft.ps1
D:\GPT4All_Local\proposed_changes\spooler job\server_response\sample_deepseek_ocr_job.pcfixe.json
D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.draft.ps1
D:\GPT4All_Local\docs\app_reference\llm_assistant\llm_assistant_app.py
```

## Compatibilite JSON app.py / spooler

Le code laptop construit le job dans `build_deepseek_ocr_job()` avec les champs suivants :

```text
job_id
type = deepseek_ocr
project_id
source_pdf
pages
output_dir
png_dir
tile_count
postprocess
retry_glitch_pages
tile_glitch_pages
fallback_tesseract_pages
```

Ces champs correspondent au contrat attendu par `run_pending_jobs.deepseek_draft.ps1`.

Compatibilite constatee :

- `type` vaut bien `deepseek_ocr` ;
- `project_id` est present ;
- `source_pdf`, `output_dir`, `png_dir` sont produits sous la racine locale PC fixe ;
- `pages` est une liste d'entiers ;
- `tile_count` vaut `2`, accepte par le spooler ;
- les quatre options sont des booleens Python serialises en booleens JSON stricts ;
- le job ne contient pas de commande arbitraire ;
- l'ecriture laptop utilise `json.dumps(..., ensure_ascii=False, indent=2)` en UTF-8.

Conclusion : le JSON produit par `app.py` est compatible avec la validation actuelle du spooler draft.

## Point d'attention contrat OCR reel

Il existe une difference a clarifier avant activation reelle :

- le dry-run laptop `build_deepseek_batch_payload()` prepare un payload avec `image_paths` ;
- le spooler draft simule actuellement un payload avec `input_dir`, `output_dir` et `pages`.

Avant d'appeler reellement `/ocr_deepseek_batch`, il faut figer un seul contrat de payload cote Flask/spooler.

Autre point a valider : le laptop prevoit actuellement des noms PNG sequentiels `page_0001.png`, `page_0002.png`, etc. meme si la page utilisateur est `8`. Il faut confirmer si la conversion reelle doit nommer les fichiers selon l'ordre extrait ou selon le numero PDF original.

## Etat de la queue avant test

Verification de :

```text
C:\Affaires\_jobs\queued
C:\Affaires\_jobs\running
C:\Affaires\_jobs\done
C:\Affaires\_jobs\failed
C:\Affaires\_jobs\logs
```

Resultat :

- `queued` : aucun JSON present ;
- `running` : aucun JSON present ;
- `done` : uniquement anciens jobs de test ASR/annotation du 2026-06-24 ;
- `failed` : uniquement anciens jobs annotation du 2026-06-24 ;
- `logs` : aucun log recent DeepSeekOCR.

Une attente d'environ 60 secondes sur `C:\Affaires\_jobs\queued` n'a detecte aucun nouveau job.

## Commande lancee

Commande executee en simulation :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_draft.ps1" -JobsRoot "C:\Affaires\_jobs" -Simulate
```

Sortie :

```text
Aucun job en attente.
```

## Resultat simulation sur queue reelle

Le spooler draft a demarre correctement et a rendu la main proprement.

Resultat :

- job traite : non, aucun job `deepseek_ocr` n'etait present dans `queued` ;
- transition `queued -> running -> done` : non observable faute de job ;
- `failed` : aucun nouveau job ajoute pendant ce test ;
- logs DeepSeekOCR : aucun nouveau log, car aucun job n'a ete pris ;
- manifest de simulation : non cree, car aucun job n'a ete pris ;
- appel HTTP : aucun ;
- conversion PNG : aucune ;
- ecriture OCR reelle : aucune.

## Validation syntaxique

Validation syntaxique PowerShell realisee :

```text
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_draft.ps1
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.draft.ps1
```

## Corrections necessaires avant activation reelle

Avant activation hors simulation :

1. Deposer un vrai JSON laptop dans `C:\Affaires\_jobs\queued` puis relancer exactement le meme test avec `-Simulate`.
2. Figer le contrat de payload `/ocr_deepseek_batch` : `image_paths` ou `input_dir + pages`.
3. Valider la convention de nommage PNG pour pages guidees, notamment page utilisateur `8`.
4. Ajouter la conversion PDF vers PNG dans un helper borne aux racines autorisees.
5. Ajouter la fusion reelle `.txt` / `.md` avec ordre des pages documente.
6. Conserver l'absence de commande libre dans les JSON.
7. Ne copier vers `D:\GPT4all_local\scripts\jobs` qu'apres validation explicite.

## Conclusion

Le contrat JSON produit par le laptop est compatible avec le spooler draft.

Le test de traitement reel en simulation n'a pas pu aller jusqu'a `done`, car aucun job laptop n'etait present dans `C:\Affaires\_jobs\queued` au moment du test. Le spooler draft a toutefois ete lance sur la queue reelle avec `-Simulate` et a confirme proprement l'absence de job en attente.

## Relance avec job laptop reel

Relance effectuee le 2026-07-06 apres depot du job reel laptop :

```text
C:\Affaires\_jobs\queued\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.json
```

JSON lu avant lancement :

- `job_id = deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940`
- `type = deepseek_ocr`
- `project_id = 2025-J47`
- `pages = [8]`
- `tile_count = 2`
- `postprocess = true`
- `retry_glitch_pages = true`
- `tile_glitch_pages = true`
- `fallback_tesseract_pages = false`

Commande lancee :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4all_local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_draft.ps1" -JobsRoot "C:\Affaires\_jobs" -Simulate
```

Resultat de transition :

- `queued` : le job n'est plus present apres execution ;
- `running` : vide apres execution ;
- `done` : le job est present ;
- `failed` : aucun nouveau job ajoute pendant cette relance. Le dossier contenait deja d'anciens jobs annotation du 2026-06-24.

Job arrive dans :

```text
C:\Affaires\_jobs\done\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.json
```

Logs crees :

```text
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.command.txt
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.stdout.log
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.stdout.manifest.json
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.stderr.log
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.exitcode.txt
C:\Affaires\_jobs\logs\20260706_234356_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260706_233940.heartbeat.txt
```

Contenu utile des logs :

```text
exitcode = 0
heartbeat = simulated 2026-07-06T23:43:56
stderr = vide
```

Le stdout confirme :

```text
SIMULATION deepseek_ocr uniquement
Aucun appel HTTP vers http://127.0.0.1:5050/ocr_deepseek_batch
Aucune conversion PDF vers PNG lancee
```

Le manifest de simulation a ete cree et contient :

- `simulated = true` ;
- `planned_endpoint = http://127.0.0.1:5050/ocr_deepseek_batch` ;
- `planned_payload.pages = [8]` ;
- `planned_payload.tile_count = 2` ;
- les quatre options booleennes au format JSON strict ;
- les chemins finaux `.merged.txt` et `.merged.md` prevus.

Verification absence d'ecriture OCR reelle :

```text
ABSENT C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\102 Assignation délivrée à 3F_pages_user_8
ABSENT C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8
```

Conclusion de la relance : le job laptop reel est compatible avec le spooler draft et passe correctement en simulation jusqu'a `done`, sans appel Flask, sans conversion PNG et sans sortie OCR reelle.
