# Analyse UI ASR Voxtral: laptop vs PC fixe

## Perimetre

Cette note compare:

- l'UI laptop de `LLM_Assistant`
- les scripts de reference PC fixe `run_transcription.bat` et `transcrire_local_voxtral.ps1`
- l'arborescence canonique affaire/captation
- le role de `infos_projet.json`

Objectif: verifier si l'UI laptop pilote correctement l'ASR Voxtral, puis proposer la cible la plus simple et la plus robuste sans casser le workflow actuel du PC fixe.

## 1. Flux actuel cote laptop

### Ce que fait `app.py`

L'ecran `Voxtral (ASR / CR)` prepare aujourd'hui des appels HTTP vers:

- `/asr_models`
- `/asr_voxtral`
- `/voxtral_chat`
- `/upload_file`

Le payload `/asr_voxtral` est prepare dans trois cas:

1. Transcription d'un media depose depuis le laptop
2. Transcription d'un media deja visible cote PC fixe
3. Generation d'un compte-rendu en passant par `/asr_voxtral`

### Champs envoyes aujourd'hui par l'UI

Selon le mode, l'UI envoie tout ou partie de:

- `audio_path`
- `model_key`
- `timestamps`
- `lang`
- `chunk`
- `stride`
- `temperature`
- `project_id`
- `output_csv_dir`
- `vocab_hint`
- `glossary_path`
- `speaker_rules_path`
- `auto_chunk`
- `batch_size`
- `export_raw_csv`
- `export_photo_csv`
- `export_srt`
- `export_vtt`
- `export_chat_csv`
- `export_chat_docx`
- `excel_encoding`
- `excel_decimal`
- `silence_split`
- `silence_top_db`
- `silence_min_ms`
- `diarize`
- `hf_token`
- `diar_options.max_speakers`
- `diar_options.min_speaker_duration`
- `diar_options.collar`
- `diar_options.allow_overlap`
- `report_prompts_path`
- `report_prompt`

### Valeurs derivees automatiquement avant correction

L'UI derivait deja automatiquement:

- `project_id`
- la liste des modeles via `/asr_models`
- des options d'export CSV/SRT/VTT
- les options de diarisation si activees

En revanche, elle ne derivait pas automatiquement a partir de `infos_projet.json`:

- le bon `.wav` de la captation
- le dossier canonique `AF_Expert_ASR/transcriptions/{id_captation}`
- le fichier canonique `id_affaire_proper_names.txt`
- le `boost_vocab.txt`

## 2. Reference PC fixe

### `run_transcription.bat`

Le batch PC fixe prend `infos_projet.json` comme point d'entree reel.

Il en deduit notamment:

- `audio` depuis `fichier_audio_source`
- `outDir` depuis `fichier_transcription`
- `boost` depuis `boost_file` ou fallback `D:\GPT4All_Local\config\boost_vocab.txt`
- `pn` depuis `proper_names_file` ou fallback `proper_names.txt` dans le dossier de sortie
- `max_speakers` depuis le JSON ou un override CLI

Puis il appelle `transcrire_local_voxtral.ps1`.

### `transcrire_local_voxtral.ps1`

Le script expose reellement:

- `LocalAudioPath`
- `ServerDropDir`
- `ProperNamesFile`
- `ModelKey`
- `Diarize`
- `MaxSpeakers`
- `Subtitles`
- `ChatDocx`
- `MinSpeakerDuration`
- `AllowOverlap`
- `Collar`
- `Boost`
- `BoostFile`
- `LocalOutDir`

Mais le POST final vers `/asr_voxtral` n'envoie pas ces noms tels quels. Il transforme surtout cela en:

- `audio_path`
- `model_key`
- `diarize`
- `diar_options`
- `output_csv_dir`
- `vocab_hint`
- `export_raw_csv`
- `export_photo_csv`
- `export_srt`
- `export_vtt`
- `excel_encoding`
- `excel_decimal`
- `auto_chunk`
- `chunk`
- `stride`
- `batch_size`

Conclusion importante:

- `ProperNamesFile` et `BoostFile` sont des parametres du script PC fixe
- cote Flask, le vrai contrat utile est plutot `vocab_hint`
- `LocalOutDir` n'est pas un parametre Flask; c'est une notion locale au script
- `ServerDropDir` correspond fonctionnellement a `output_csv_dir`

## 3. Ecarts identifies avant correction

### Ce que l'UI faisait deja correctement

- choix de `ModelKey`
- activation de `Diarize`
- reglage de `MaxSpeakers`
- appel direct a `/asr_voxtral`
- passage de `project_id`
- export CSV/SRT/VTT

### Ce qui manquait

- aucun usage de `infos_projet.json` dans `app.py`
- aucun ciblage automatique de `AE_Expert_captations/{id_captation}/audio`
- aucun ciblage automatique de `AF_Expert_ASR/transcriptions/{id_captation}`
- pas de proposition automatique du fichier `id_affaire_proper_names.txt`
- pas de prise en compte automatique du `boost_vocab.txt`
- le fallback de sortie serveur restait trop generique

### Point bloquant principal

La route serveur `_resolve_asr_out_dir()` resout, si `output_csv_dir` est absent:

1. `output_csv_dir` explicite
2. `project_config.paths.asr_transcriptions`
3. le dossier parent de `audio_path`

Or `project_config.paths.asr_transcriptions` pointe sur `AF_Expert_ASR/transcriptions` sans `id_captation`.

Donc, sans surcharge explicite cote client, les sorties ne sont pas garanties dans:

`AF_Expert_ASR/transcriptions/{id_captation}`

## 4. Integration avec `infos_projet.json`

### Constat

Le repo contient deja `tools/sync/write_infos_projet.py`, qui ecrit un `infos_projet.json` avec:

- `id_affaire`
- `id_captation`
- `fichier_audio_source`
- `pcfixe.fichier_audio_source`
- `pcfixe.fichier_transcription`
- `pcfixe.boost_file`
- `pcfixe.proper_names_file`
- `pcfixe.max_speakers`

Ce fichier etait donc deja la bonne base de derivation metier, mais l'UI ne le consultait pas.

### Correction appliquee

L'UI Voxtral a ete adaptee pour:

- lister les captations detectees pour l'affaire
- resoudre le contexte d'une captation selectionnee
- lire `infos_projet.json` si present
- deduire le media cote PC fixe
- deduire le dossier de sortie canonique `AF_Expert_ASR/transcriptions/{id_captation}`
- viser le fichier canonique `AF_Expert_ASR/transcriptions/{id_captation}/{id_affaire}_proper_names.txt`
- chercher `boost_vocab.txt` en priorite dans les emplacements connus
- fusionner automatiquement le contenu de `proper_names` et `boost_vocab` dans `vocab_hint`

### Effet pratique

Desormais, quand une captation est selectionnee, l'UI peut piloter `/asr_voxtral` avec:

- `audio_path` deduit
- `output_csv_dir` force vers le dossier canonique de la captation
- `vocab_hint` enrichi automatiquement

Cela rapproche le fonctionnement laptop du workflow reel du PC fixe sans imposer de refonte serveur.

## 5. Comparaison A / B / C

### A. Mode actuel: scripts locaux PC fixe

Avantages:

- reference fonctionnelle existante
- robuste pour les chemins Windows locaux
- deja aligne sur `infos_projet.json`

Risques:

- peu pilotable depuis l'UI laptop
- logique de parametrage repartie entre batch, PowerShell et Flask
- moins lisible pour l'utilisateur final cote laptop

Impact pipeline:

- nul, c'est l'existant

Compatibilite NAS / synchro:

- bonne, car tout est pilote depuis le PC fixe

### B. Pilotage depuis laptop vers l'API Flask

Avantages:

- UI plus simple a utiliser
- pas de duplication lourde de logique serveur
- garde `/asr_voxtral` comme point d'execution
- permet de forcer explicitement le bon `output_csv_dir`

Risques:

- si le client derive mal les chemins, il peut sortir du canon
- il faut rester strict sur l'arborescence et les chemins UNC/PC fixe

Impact pipeline:

- faible si le laptop ne fait que resoudre les chemins puis appeler l'API

Compatibilite NAS / synchro:

- bonne si le point de verite des chemins reste le dossier affaire et `infos_projet.json`

### C. Execution via serveur Flask avec encapsulation du script

Deux variantes:

- route Flask qui declenche `transcrire_local_voxtral.ps1`
- integration directe de la logique du script dans Flask

Avantages:

- centralisation maximale
- moins de logique cote client

Risques:

- augmentation du couplage serveur
- risque de doublonner ce qui existe deja entre PowerShell et Flask
- maintenance plus delicate si la logique diverge

Impact pipeline:

- moyen a fort selon la variante

Compatibilite NAS / synchro:

- potentiellement bonne, mais plus intrusive

## 6. Cible recommandee

### Recommandation

La cible la plus simple et la plus robuste est:

- garder le PC fixe / Flask comme moteur d'execution
- garder les scripts PC fixe comme reference et filet de securite
- faire du laptop un client intelligent mais leger
- resoudre cote laptop les chemins canoniques depuis `id_affaire`, `id_captation` et `infos_projet.json`
- envoyer explicitement `audio_path`, `output_csv_dir` et `vocab_hint` a `/asr_voxtral`

### Pourquoi cette cible

- elle ne casse pas le batch PC fixe
- elle reste compatible avec l'existant Flask
- elle evite de repliquer la logique ASR profonde dans le laptop
- elle limite les problemes de synchronisation NAS, car les sorties sont forcees au bon endroit

## 7. Point de verite recommande

### Pour les chemins fichiers

Point de verite recommande:

- arborescence canonique affaire/captation
- `infos_projet.json` quand il existe
- `roots.pcfixe` + chemins canoniques pour construire les chemins absolus cote serveur

L'UI laptop ne doit pas inventer une autre logique de rangement.

### Pour les parametres ASR

Point de verite recommande:

- contrat Flask `/asr_voxtral` pour les parametres d'execution reels
- scripts PC fixe comme reference fonctionnelle de transformation

Traduction pratique:

- `ModelKey`, `Diarize`, `MaxSpeakers` restent de vrais parametres UI
- `ProperNamesFile` et `BoostFile` sont surtout des sources de `vocab_hint`
- `ServerDropDir` se traduit par `output_csv_dir`
- `LocalOutDir` n'a pas a devenir un parametre central de l'UI laptop

## 8. Diff applique

### Fichier modifie

- `C:\CodexWorkspace\LLM_Assistant\app.py`

### Changement minimal applique

- ajout d'un resolveur de contexte ASR par captation
- lecture de `infos_projet.json` si present
- derivation du media PC fixe et du dossier de sortie canonique
- auto-injection du vocabulaire issu de `proper_names` et `boost_vocab`
- forçage de `output_csv_dir` vers `AF_Expert_ASR/transcriptions/{id_captation}` quand une captation est selectionnee
- adaptation des miroirs de sortie pour lire le bon dossier cible

## 9. Conclusion courte

### Ce que l'UI faisait deja bien

- appeler correctement `/asr_voxtral`
- exposer `ModelKey`, `Diarize`, `MaxSpeakers`
- permettre les exports utiles

### Ce qui manquait

- l'usage de `infos_projet.json`
- le ciblage automatique de la captation
- le bon dossier de sortie par `id_captation`
- le chainage automatique de `proper_names` et `boost_vocab`

### Ce qui doit rester le point de verite

- chemins: arborescence canonique + `infos_projet.json` + `roots.pcfixe`
- parametres ASR reels: contrat Flask `/asr_voxtral`
- reference fonctionnelle: scripts du PC fixe
