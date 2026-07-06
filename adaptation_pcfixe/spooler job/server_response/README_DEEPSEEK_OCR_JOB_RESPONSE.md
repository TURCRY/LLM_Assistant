# Reponse PC fixe - job deepseek_ocr draft

## Portee

Cette reponse prepare uniquement une proposition d'ajout du type de job `deepseek_ocr`.

Aucun fichier de production n'a ete modifie :

```text
D:\GPT4all_local\scripts\jobs
```

Aucun appel DeepSeekOCR reel n'a ete lance, aucun appel HTTP vers Flask n'a ete effectue, et Flask / `app.py` ne sont pas modifies.

## Audit du spooler existant

Base analysee :

```text
README_PC_FIXE_RESPONSE.md
run_pending_jobs.pcfixe.ps1
sample_asr_job.pcfixe.json
sample_annotation_job.pcfixe.json
```

Le spooler PC fixe actuel est compatible avec l'ajout local d'un troisieme type de job :

- il cree les dossiers `queued`, `running`, `done`, `failed`, `logs` ;
- il prend un seul job JSON a la fois ;
- il utilise un verrou `spooler.lock` avec detection de verrou perime ;
- il conserve la reprise `running -> failed` par defaut via `-RecoverRunning fail` ;
- il refuse les types inconnus ;
- il reconstruit lui-meme les commandes autorisees pour `asr_voxtral` et `annotation_photos_batch` ;
- il journalise stdout, stderr, exitcode, heartbeat et commande.

Le point d'extension naturel est donc le `switch ([string]$job.type)`.

## Format JSON propose

Exemple fourni :

```text
sample_deepseek_ocr_job.pcfixe.json
```

Champs proposes :

```json
{
  "job_id": "deepseek_ocr_2025-J47_102_assignation_page_8_20260706_120000",
  "type": "deepseek_ocr",
  "project_id": "2025-J47",
  "source_pdf": "C:\\Affaires\\2025-J47\\AA_Expert_Admin\\Depot_initial\\_Guided_Analysis\\102 Assignation délivrée à 3F_pages_user_8.pdf",
  "pages": [8],
  "output_dir": "C:\\Affaires\\2025-J47\\AD_Expert_Traitements\\_OCR_Dire_Bordereau\\102 Assignation délivrée à 3F_pages_user_8",
  "png_dir": "C:\\Affaires\\2025-J47\\AD_Expert_Traitements\\_OCR_Dire_Bordereau\\_pages_png\\102 Assignation délivrée à 3F_pages_user_8",
  "tile_count": 2,
  "postprocess": true,
  "retry_glitch_pages": true,
  "tile_glitch_pages": true,
  "fallback_tesseract_pages": false
}
```

Validations proposees cote spooler :

- `type` doit valoir `deepseek_ocr` ;
- `project_id` doit suivre le format `2025-J47` ;
- `source_pdf`, `output_dir`, `png_dir` doivent rester sous `C:\Affaires` ou `D:\GPT4all_local` ;
- `source_pdf` doit avoir l'extension `.pdf` ;
- `pages` doit contenir des entiers strictement positifs, sans doublon ;
- `tile_count` est limite a `2` ou `3` ;
- `postprocess`, `retry_glitch_pages`, `tile_glitch_pages`, `fallback_tesseract_pages` doivent etre des booleens JSON stricts ;
- aucune commande libre n'est acceptee dans le JSON.

## Responsabilites proposees

### app.py laptop

Responsabilites futures recommandees :

- preparer le PDF guide et le choix des pages ;
- generer le `job_id` ;
- ecrire le JSON dans la queue du spooler PC fixe ;
- suivre le statut dans `queued`, `running`, `done`, `failed` ;
- lire les logs et les chemins finaux produits par le PC fixe.

`app.py` ne devrait pas fournir de commande PowerShell ou Python arbitraire.

### Spooler PC fixe

Responsabilites futures recommandees :

- valider strictement le JSON ;
- convertir les pages PDF demandees en PNG dans `png_dir` ;
- construire un payload connu pour Flask local ;
- appeler uniquement `http://127.0.0.1:5050/ocr_deepseek_batch` ;
- fusionner les sorties `.txt` et `.md` ;
- ecrire les chemins finaux attendus ;
- deplacer le job vers `done` ou `failed`.

Dans le draft fourni, cette branche est volontairement en simulation uniquement.

### Flask PC fixe

Responsabilites futures recommandees :

- conserver la logique DeepSeekOCR dans la route locale existante ou prevue ;
- recevoir un payload borne et explicite ;
- produire les fichiers OCR par page ;
- ne pas connaitre directement la file de jobs.

Pour ce lot, Flask n'est pas modifie.

## Fichiers fournis

```text
README_DEEPSEEK_OCR_JOB_RESPONSE.md
run_pending_jobs.deepseek_draft.ps1
sample_deepseek_ocr_job.pcfixe.json
run_deepseek_ocr_job.draft.ps1
```

`run_pending_jobs.deepseek_draft.ps1` conserve les branches existantes :

- `asr_voxtral` ;
- `annotation_photos_batch`.

Il ajoute :

- `deepseek_ocr`, simulation uniquement ;
- manifest de simulation dans les logs du spooler ;
- aucun appel HTTP ;
- aucune conversion PDF ;
- aucune ecriture sous `C:\Affaires` pendant la simulation.

`run_deepseek_ocr_job.draft.ps1` est un helper optionnel de simulation. Il valide un JSON `deepseek_ocr` et ecrit un plan de traitement a cote du JSON fourni.

## Tests realises le 2026-07-06

Validation syntaxique PowerShell :

```text
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_pending_jobs.deepseek_draft.ps1
PARSE_OK D:\GPT4All_Local\proposed_changes\spooler job\server_response\run_deepseek_ocr_job.draft.ps1
```

Simulation spooler avec le sample `deepseek_ocr` :

```text
JobsRoot:
D:\GPT4All_Local\proposed_changes\spooler job\server_response\_deepseek_spooler_test

Resultat:
DONE_COUNT=1
FAILED_COUNT=0
```

Le stdout de simulation indique :

```text
SIMULATION deepseek_ocr uniquement
Aucun appel HTTP vers http://127.0.0.1:5050/ocr_deepseek_batch
Aucune conversion PDF vers PNG lancee
```

Le manifest de simulation contient bien :

- `planned_endpoint = http://127.0.0.1:5050/ocr_deepseek_batch` ;
- `planned_payload.pages = [8]` ;
- `tile_count = 2` ;
- les quatre options booleennes au format JSON strict ;
- les chemins finaux `.merged.txt` et `.merged.md` prevus.

## Risques

- Le contrat exact de `/ocr_deepseek_batch` doit etre confirme avant de passer de la simulation a l'execution reelle.
- La conversion PDF vers PNG doit choisir un outil present sur le PC fixe et teste avec chemins Windows accentues.
- La fusion `.txt` / `.md` doit definir l'ordre exact des pages et le comportement si une page OCR echoue.
- Le spooler ne doit pas relancer automatiquement un job DeepSeekOCR partiel sans verification des sorties deja produites.
- Les chemins avec accents et espaces doivent rester testes de bout en bout.

## Lots suivants proposes

1. Valider le contrat JSON avec le laptop et le payload exact attendu par `/ocr_deepseek_batch`.
2. Ajouter un helper reel de conversion PDF vers PNG, borne aux racines autorisees.
3. Tester localement `/ocr_deepseek_batch` avec un dossier PNG minimal, hors spooler.
4. Ajouter la fusion `.txt` / `.md` dans le helper reel.
5. Integrer le helper reel au spooler, avec `deepseek_ocr` toujours sans commande libre.
6. Seulement apres validation, recopier la version retenue vers `D:\GPT4all_local\scripts\jobs`.

## A recopier vers le laptop

Pour revue laptop, recopier :

```text
README_DEEPSEEK_OCR_JOB_RESPONSE.md
run_pending_jobs.deepseek_draft.ps1
sample_deepseek_ocr_job.pcfixe.json
run_deepseek_ocr_job.draft.ps1
```
