# Lot qualite DeepSeekOCR bordereau

Date: 2026-07-07

## Perimetre

- Script modifie: `D:\GPT4All_Local\scripts\jobs\run_deepseek_ocr_job.ps1`
- Flask non modifie.
- `app.py` non modifie.
- Tache planifiee non modifiee.
- Aucun fallback Tesseract automatique ajoute.

## Audit du job de reference

Job audite:

`deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_20260707_015812`

Dossier OCR:

`C:\Affaires\2025-J47\AD_Expert_Traitements\_OCR_Dire_Bordereau\102 Assignation délivrée à 3F_pages_user_8`

Constats:

- `raw.txt`, `clean.txt`, `final.md` et `final.txt` contiennent la ligne parasite initiale `Pro criticality`.
- `raw.txt` et `clean.txt` sont identiques sur ce cas: le postprocess legal OCR n'a pas introduit cette ligne.
- La reponse batch indique `postprocess=true`, `glitch=false`, `retry_triggered=false`, `tile_triggered=false`, `fallback_tesseract_triggered=false`.
- Le prompt demande deja une transcription stricte sans correction, inference ni reformulation.
- Le PNG source commence visuellement par `BORDEREAU DE PIECES`; la mention `Pro criticality` n'est pas visible sur l'image.
- Le PNG source contient les pieces 01 a 06, dont la piece 04.
- La sortie DeepSeekOCR contient les pieces 01, 02, 03, 05 et 06, mais omet la piece 04.

Conclusion audit:

- `Pro criticality` provient de la sortie brute DeepSeekOCR, pas de la fusion finale ni du postprocess.
- L'omission de la piece 04 est une erreur OCR residuelle. Elle ne doit pas etre masquee par le nettoyage.

## Correction appliquee

La correction est limitee a la phase de fusion dans `run_deepseek_ocr_job.ps1`.

Ajouts:

- Nettoyage prudent avant fusion:
  - recherche d'une ligne contenant `BORDEREAU DE PIECES`;
  - suppression uniquement des lignes avant ce titre;
  - conservation du titre et du reste du texte OCR.
- Journalisation dans le manifest final:
  - `leading_noise_removed`
  - `leading_noise_removed_text`
- Controle qualite bordereau:
  - detection des numeros `Piece n XX` / `Piece n° XX`;
  - detection des ruptures de sequence;
  - journalisation:
    - `bordereau_piece_numbers_detected`
    - `bordereau_missing_piece_numbers`
    - `quality_warning`

## Test de validation

Job de test:

`deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty`

Commande lancee manuellement:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\GPT4All_Local\scripts\jobs\run_pending_jobs.ps1" -JobsRoot "C:\Affaires\_jobs"
```

Resultat spooler:

- `queued`: false
- `running`: false
- `done`: true
- `failed`: false
- exitcode: 0
- stderr: vide

Logs crees:

- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.command.txt`
- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.exitcode.txt`
- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.heartbeat.runner.cmd`
- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.heartbeat.txt`
- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.stderr.log`
- `C:\Affaires\_jobs\logs\20260707_023551_deepseek_ocr_2025-J47_102_Assignation_delivree_a_3F_pages_user_8_quality_20260707_qlty.stdout.log`

Artefacts verifies:

- `final.md`: cree, ne contient plus `Pro criticality`.
- `final.txt`: cree, ne contient plus `Pro criticality`.
- `final_manifest.json`: cree.
- `ocr_deepseek_batch.response.json`: cree.
- `ocr_deepseek_batch.diagnostics.json`: cree.

Champs qualite constates dans le manifest:

```json
{
  "leading_noise_removed": true,
  "leading_noise_removed_text": "Pro criticality",
  "bordereau_piece_numbers_detected": [1, 2, 3, 5, 6],
  "bordereau_missing_piece_numbers": [4],
  "quality_warning": true
}
```

Extrait final verifie:

```text
## Page 8

# **BORDEREAU DE PIECES**

|Piece n 01 :|Attestation de propriete des epoux BROCQ.|
...
|Piece n 05 :|Photographies des desordres affectant la maison 8 rue des Carrieres a 95300 PONTOISE.|
|Piece n 06 :|Rapport d'expertise de Monsieur l'Expert Laurent PIRES, Cabinet ELEX du 18 decembre 2024.|
```

## Limites restantes

- La correction supprime seulement le bruit place avant le titre `BORDEREAU DE PIECES`.
- Elle ne corrige pas l'omission de la piece 04.
- Le fallback Tesseract n'est pas relance dans ce lot.
- Le controle qualite signale les ruptures mais ne bloque pas le job: le job reste `done` avec `quality_warning=true`.

## Suite recommandee

- Lot suivant: utiliser `quality_warning=true` pour declencher une verification humaine ou un fallback cible, sans boucle automatique non controlee.
