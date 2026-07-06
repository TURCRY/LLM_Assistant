# Integration production controlee - deepseek_ocr

Date : 2026-07-07

## Portee

Integration controlee du support `deepseek_ocr` dans les scripts de production :

```text
D:\GPT4All_Local\scripts\jobs
```

Contraintes respectees :

- `app.py` non modifie ;
- Flask non modifie ;
- tache planifiee non modifiee ;
- integration limitee aux scripts jobs demandes ;
- test manuel uniquement.

## Fichiers production modifies / ajoutes

Fichier modifie :

```text
D:\GPT4All_Local\scripts\jobs\run_pending_jobs.ps1
```

Fichier ajoute :

```text
D:\GPT4All_Local\scripts\jobs\run_deepseek_ocr_job.ps1
```

Les jobs existants sont conserves :

- `asr_voxtral`
- `annotation_photos_batch`

Le job ajoute est :

- `deepseek_ocr`

## Comportement integre

Le flux `deepseek_ocr` de production reprend le comportement valide au lot 9C :

1. validation stricte du job JSON ;
2. validation des chemins sous `C:\Affaires` ou `D:\GPT4all_local` ;
3. conversion PDF vers PNG avec PyMuPDF via :

```text
D:\GPT4all_local\.venv\Scripts\python.exe
```

4. appel borne a :

```text
http://127.0.0.1:5050/ocr_deepseek_batch
```

5. ecriture de la reponse JSON et des diagnostics ;
6. fusion finale en `.final.md` et `.final.txt` ;
7. ecriture d'un `.final_manifest.json` ;
8. deplacement du job vers `done` ou `failed` par le spooler.

Le helper production est autonome : il ne depend pas des scripts de proposition `server_response`.

## Gardes-fous conserves

- un seul job traite a la fois ;
- verrou `spooler.lock` ;
- recovery `running -> failed` conserve ;
- aucune commande arbitraire depuis le JSON ;
- endpoint DeepSeekOCR force a `127.0.0.1:5050` ;
- cle API lue depuis la configuration locale, non ecrite dans les logs ;
- `tile_count = 2` pour cette integration initiale ;
- booleens stricts valides ;
- integration initiale limitee a une seule page, comme le test valide.

## Validations syntaxiques

```text
PARSE_OK D:\GPT4All_Local\scripts\jobs\run_pending_jobs.ps1
PARSE_OK D:\GPT4All_Local\scripts\jobs\run_deepseek_ocr_job.ps1
```

## Test production manuel

Job cree en queue :

```text
C:\Affaires\_jobs\queued\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.json
```

Commande lancee manuellement :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4All_Local\scripts\jobs\run_pending_jobs.ps1" -JobsRoot "C:\Affaires\_jobs"
```

Resultat spooler :

- `queued` : job absent apres execution ;
- `running` : job absent apres execution ;
- `done` : job present ;
- `failed` : job absent ;
- exit code : `0` ;
- stderr : vide.

Job final :

```text
C:\Affaires\_jobs\done\deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.json
```

Logs :

```text
C:\Affaires\_jobs\logs\20260707_004800_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.command.txt
C:\Affaires\_jobs\logs\20260707_004800_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.stdout.log
C:\Affaires\_jobs\logs\20260707_004800_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.stderr.log
C:\Affaires\_jobs\logs\20260707_004800_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.exitcode.txt
C:\Affaires\_jobs\logs\20260707_004800_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.heartbeat.txt
```

## Sorties finales du test

Dossier OCR :

```text
C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8
```

Fichiers produits :

```text
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.final.md
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.final.txt
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.final_manifest.json
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.ocr_deepseek_batch.response.json
deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_prod_20260707_010000.ocr_deepseek_batch.diagnostics.json
```

Le manifest final indique :

```text
project_id = 2025-J47
pages = [8]
selected_field = clean_text_path
text_chars = 562
metrics.count = 1
metrics.total_text_chars = 554
```

## Verification contenu

Controle lexical dans le fichier final Markdown :

```text
Pièce n° 01 : presente
Pièce n° 02 : presente
Pièce n° 03 : presente
Pièce n° 04 : absente
Pièce n° 05 : presente
Pièce n° 06 : presente
```

Conclusion : l'integration production fonctionne techniquement. L'absence de `Pièce n° 04` est un resultat OCR deja observe au lot 9C et doit etre verifie fonctionnellement sur le document source.

## Limites connues

- Integration initiale limitee a une seule page.
- Pas encore de validation multi-page.
- Pas de modification de la tache planifiee.
- Pas de correction automatique du contenu OCR.
- Les PNG existants peuvent etre reecrits si le meme dossier `png_dir` est reutilise.

## Validation

Test production manuel : OK.

Commit Git dedie realise apres test OK :

```text
741521f Add deepseek OCR job spooler support
```
