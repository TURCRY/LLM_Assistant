# Schéma `infos_projet.json`

`infos_projet.json` est le contrat de captation partagé entre les briques du pipeline :

- `run_all_from_jpg_v5.bat` ;
- `selection_fichiers_interface.py` ;
- `LLM_Assistant` ;
- `AnnotationPhotosGPT` ;
- spooler ASR ;
- pipeline compte-rendu ;
- serveur `GPT4All_Local`.

La source de vérité fonctionnelle est le fichier `infos_projet.json` publié dans l’arborescence canonique de l’affaire / captation, notamment :

```text
<racine Affaires>/<id_affaire>/AF_Expert_ASR/transcriptions/<id_captation>/infos_projet.json
```

Les applications clientes peuvent en conserver une copie locale ou remappée, mais ne doivent pas créer de doctrine concurrente de chemins métier.

## Principes

1. `infos_projet.json` décrit une captation d’une affaire.
2. Les chemins métier pointent vers les dossiers canoniques de l’affaire/captation.
3. Les chemins temporaires de traitement ne sont jamais la source de vérité.
4. Les blocs `pcfixe`, `nas`, `roots` et `debrief` peuvent fournir des chemins remappés selon le contexte d’exécution.
5. Les consommateurs doivent accepter les clés historiques tant qu’une migration explicite n’a pas été faite.

## Clés minimales obligatoires

| Clé | Type | Rôle |
|---|---:|---|
| `id_affaire` | string | Identifiant canonique de l’affaire, ex. `2025-J47`. |
| `id_captation` | string | Identifiant canonique de captation, ex. `accedit-2025-11-13`. |

Aliases historiques tolérés :

- `project_id` pour `id_affaire` ;
- `captation_id` pour `id_captation`.

Ces aliases peuvent être lus pour compatibilité, mais les nouveaux producteurs doivent écrire `id_affaire` et `id_captation`.

## Racines et contextes de chemins

### `roots`

Bloc optionnel recommandé :

```json
{
  "roots": {
    "laptop": "C:\\Affaires\\2025-J47",
    "pcfixe": "\\\\192.168.0.155\\Affaires\\2025-J47",
    "nas": "\\\\192.168.1.20\\Affaires\\2025-J47"
  }
}
```

### `pcfixe`

Bloc de chemins vus par le PC fixe ou par le serveur :

```json
{
  "pcfixe": {
    "root_affaires": "C:\\Affaires",
    "infos": "C:\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\infos_projet.json",
    "fichier_transcription": "C:\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\transcription.csv",
    "fichier_photos": "C:\\Affaires\\2025-J47\\AE_Expert_captations\\accedit-2025-11-13\\photos\\photos.csv",
    "fichier_photos_batch": "C:\\Affaires\\2025-J47\\AE_Expert_captations\\accedit-2025-11-13\\photos\\photos_batch.csv"
  }
}
```

Le champ `pcfixe.root_affaires` est important pour le pipeline compte-rendu lorsqu’il doit reconstruire :

```text
<root_affaires>/<id_affaire>/BE_Traitement_captations/<id_captation>/compte_rendu_LLM
```

## Bloc photos

Clés historiques de premier niveau :

| Clé | Type | Rôle |
|---|---:|---|
| `fichier_photos` | string | Chemin de `photos.csv`. |
| `fichier_photos_batch` | string | Chemin de `photos_batch.csv`, produit par le flux JPG. |

Chemins remappés acceptés :

```json
{
  "pcfixe": {
    "fichier_photos": "C:\\Affaires\\...\\photos\\photos.csv",
    "fichier_photos_batch": "C:\\Affaires\\...\\photos\\photos_batch.csv"
  }
}
```

Règles :

- `photos.csv` et `photos_batch.csv` doivent rester dans le dossier métier de captation :

```text
AE_Expert_captations/<id_captation>/photos/
```

- Aucun producteur ne doit proposer un dossier applicatif comme dossier métier, notamment :
  - `C:\AnnotationPhotosGPT\batch_pcfixe`
  - `C:\AnnotationPhotosGPT\data\batch_cache`

## Bloc ASR

Clés historiques de premier niveau :

| Clé | Type | Rôle |
|---|---:|---|
| `fichier_audio_source` | string | WAV natif ou source audio captation. |
| `audio_compat_source` | string | Source audio compatible ou remappée. |
| `fichier_audio_compatible` | string | WAV préparé compatible ASR si déjà produit. |
| `fichier_transcription` | string | CSV principal de transcription ASR. |
| `fichier_contexte_general` | string | Contexte JSON associé, souvent à côté du CSV. |

Chemins remappés côté PC fixe :

```json
{
  "pcfixe": {
    "fichier_audio_source": "C:\\Affaires\\...\\audio\\captation.wav",
    "fichier_audio_compatible": "C:\\Affaires\\...\\audio\\captation_mono16_16000Hz.wav",
    "fichier_transcription": "C:\\Affaires\\...\\AF_Expert_ASR\\transcriptions\\<id_captation>\\transcription.csv",
    "fichier_contexte_general": "C:\\Affaires\\...\\AF_Expert_ASR\\transcriptions\\<id_captation>\\contexte_general_compte_rendu.json"
  }
}
```

Le spooler ASR `schema_version=2` reçoit au minimum :

```json
{
  "schema_version": 2,
  "job_id": "...",
  "type": "asr_voxtral",
  "infos_projet": "C:\\Affaires\\...\\infos_projet.json",
  "audio_input": "C:\\Affaires\\...\\audio.wav",
  "audio_prepared": "C:\\Affaires\\...\\audio_mono16_16000Hz.wav",
  "proper_names_file": "",
  "boost_file": "D:\\GPT4All_Local\\config\\boost_vocab.txt",
  "debrief_csv": "",
  "model": "Voxtral_Mini_3B_Transformers",
  "diarize": true
}
```

## Bloc debrief

Le debrief audio est optionnel et distinct de l’audio principal.

Bloc recommandé :

```json
{
  "debrief": {
    "enabled": true,
    "type": "audio",
    "source_laptop": "C:\\...\\debrief\\debrief.wav",
    "pcfixe_wav": "C:\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\debrief\\debrief.wav",
    "nas_wav": "\\\\192.168.1.20\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\debrief\\debrief.wav",
    "csv": "",
    "transcribed": false,
    "proper_names_file": "",
    "proper_names_source": "none"
  }
}
```

Valeurs admises pour `proper_names_source` :

- `canonical` : fichier canonique de la captation ;
- `debrief_folder` : fichier trouvé dans le dossier du WAV debrief ;
- `manual` : chemin saisi manuellement ;
- `none` : aucun fichier transmis.

Quand le debrief audio est transcrit séparément, `debrief.csv` doit être renseigné soit :

- dans `debrief.csv` ;
- ou dans une clé historique compatible avec le pipeline compte-rendu, par exemple `fichier_debrief` / `pcfixe.fichier_debrief`, tant que ce pipeline n’a pas migré vers `debrief.csv`.

## Proper names

Clés admises :

| Clé | Type | Rôle |
|---|---:|---|
| `proper_names_file` | string | Fichier de noms propres principal. |
| `pcfixe.proper_names_file` | string | Remappage côté PC fixe. |
| `debrief.proper_names_file` | string | Fichier spécifique au debrief. |

Règles :

- Le proper names principal de la captation ne doit pas être modifié par le choix du debrief.
- Le debrief peut utiliser un fichier différent ou aucun fichier.
- Les consommateurs doivent accepter `*_proper_names.txt` dans le dossier canonique de transcription.

## Compte-rendu

Bloc optionnel :

```json
{
  "compte_rendu": {
    "provider": "openai",
    "api_base": "http://openai-adapter:5055",
    "model_pass1": "annoter_segments_remote",
    "model_pass2": "annoter_segments_remote",
    "model_pass3": "annoter_segments_remote_alt",
    "preset": "equilibre",
    "pseudo_api_base": "",
    "pseudo_job_id": ""
  }
}
```

Le pipeline compte-rendu lit :

- `id_affaire` ;
- `id_captation` ;
- `fichier_transcription` ou `pcfixe.fichier_transcription` ;
- `fichier_contexte_general` ou `pcfixe.fichier_contexte_general`, avec fallback à côté du CSV ;
- `fichier_sujets` / `sujets_path` ou `pcfixe.*`, avec fallback `Sujets.xlsx` à côté du CSV ;
- `fichier_participants` / `participants_path` ou `pcfixe.*`, avec fallback `Participants.xlsx` à côté du CSV ;
- `fichier_debrief` ou `pcfixe.fichier_debrief` selon le wrapper historique.

Sortie canonique :

```text
BE_Traitement_captations/<id_captation>/compte_rendu_LLM/
```

## `manifest.json`

Chaque job long doit publier un manifeste de fin de job.

Pour l’ASR :

```json
{
  "job_type": "asr",
  "status": "completed",
  "id_affaire": "2025-J47",
  "id_captation": "accedit-2025-11-13",
  "audio_input": "...",
  "audio_prepared": "...",
  "debrief_csv": "",
  "proper_names_file": "",
  "model": "Voxtral_Mini_3B_Transformers",
  "diarize": true,
  "started_at": "...",
  "finished_at": "...",
  "duration_seconds": 0,
  "work_dir": "...",
  "outputs": [],
  "copied_to_pcfixe": [],
  "copied_to_nas": [],
  "errors": []
}
```

Emplacements attendus :

- dossier temporaire du job ;
- dossier canonique :

```text
AF_Expert_ASR/transcriptions/<id_captation>/manifest.json
```

## Compatibilité ascendante

Les consommateurs doivent :

1. Lire d’abord les clés nouvelles explicites.
2. Conserver les fallbacks historiques.
3. Ne pas supprimer les clés inconnues.
4. Ne pas réécrire tout le JSON si seule une section doit être ajoutée.
5. Ne pas renommer sans migration :
   - `id_affaire`
   - `id_captation`
   - `project_id`
   - `captation_id`
   - `fichier_photos`
   - `fichier_photos_batch`
   - `fichier_transcription`
   - `fichier_contexte_general`
   - `fichier_debrief`
   - `pcfixe`
   - `roots`
   - `debrief`
   - `compte_rendu`

## Champs à ne pas détourner

- `fichier_transcription` désigne le CSV principal ASR, pas le debrief.
- `debrief.source_laptop` désigne le WAV debrief source, pas l’audio principal.
- `debrief.csv` désigne la transcription CSV debrief lorsqu’elle existe.
- `fichier_photos_batch` désigne `photos_batch.csv` métier, jamais un cache applicatif.
- `pcfixe.root_affaires` désigne la racine Affaires vue côté PC fixe, pas un dossier temporaire.

## Exemple minimal

```json
{
  "id_affaire": "2025-J47",
  "id_captation": "accedit-2025-11-13",
  "project_id": "2025-J47",
  "captation_id": "accedit-2025-11-13",
  "fichier_photos": "C:\\...\\photos\\photos.csv",
  "fichier_photos_batch": "C:\\...\\photos\\photos_batch.csv",
  "fichier_transcription": "C:\\...\\transcriptions\\accedit-2025-11-13\\transcription.csv",
  "proper_names_file": "C:\\...\\transcriptions\\accedit-2025-11-13\\2025-J47_proper_names.txt",
  "pcfixe": {
    "root_affaires": "C:\\Affaires",
    "fichier_photos": "C:\\Affaires\\2025-J47\\AE_Expert_captations\\accedit-2025-11-13\\photos\\photos.csv",
    "fichier_photos_batch": "C:\\Affaires\\2025-J47\\AE_Expert_captations\\accedit-2025-11-13\\photos\\photos_batch.csv",
    "fichier_transcription": "C:\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\transcription.csv"
  },
  "debrief": {
    "enabled": true,
    "type": "audio",
    "source_laptop": "C:\\...\\debrief\\debrief.wav",
    "pcfixe_wav": "C:\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\debrief\\debrief.wav",
    "nas_wav": "\\\\192.168.1.20\\Affaires\\2025-J47\\AF_Expert_ASR\\transcriptions\\accedit-2025-11-13\\debrief\\debrief.wav",
    "csv": "",
    "transcribed": false,
    "proper_names_file": "",
    "proper_names_source": "none"
  }
}
```
