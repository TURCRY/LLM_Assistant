Ce fichier `.bat` est un **script d’automatisation Windows** qui sert à **préparer, compléter et copier un dossier de captation photo/audio** vers une arborescence centralisée d’affaire.

En pratique, il fait ceci :

1. **Vérifie les paramètres et le dossier courant**

   * il exige un identifiant d’affaire en argument, par exemple `2025-J46` ;
   * il doit être lancé depuis un dossier `...\Photos\JPG`.

2. **Reconstitue les chemins sources**

   * dossier des JPG natifs ;
   * dossier `Photos` parent ;
   * dossier de captation parent ;
   * dossier `Audio` frère.

3. **Déduit un identifiant de captation**

   * à partir du nom du dossier de captation, par exemple
     `Accedit 06 11 2025` → `accedit-2025-11-06`.

4. **Cherche les fichiers de travail photo**

   * détecte un CSV “UI” dans le dossier `Photos` ;
   * crée `photos_batch.csv` s’il n’existe pas.

5. **Cherche les fichiers audio/transcription**

   * le CSV de transcription ;
   * un WAV converti en mono 16 kHz ;
   * un WAV source ;
   * un fichier `contexte_general*.json`.

6. **Crée l’arborescence de destination**

   * sous un partage réseau de type :
     `\\192.168.1.20\volume1\Affaires\<AFFAIRE>\...`

7. **Met à jour les métadonnées photo**

   * appelle un script Python pour enrichir `photos.csv` avec les champs de destination.

8. **Copie les fichiers**

   * JPG natifs ;
   * JPG réduits ;
   * fichiers CSV / XLS / XLSX du dossier `Photos` ;
   * fichiers audio ;
   * transcriptions, sous-titres, JSON, TXT, XLSX du dossier `Audio`.

9. **Dépose des fichiers de configuration**

   * `prompt_gpt.json`
   * `prompt_gpt_batch_only.json`
   * `config_llm.json`

10. **Copie certains annexes de contexte**

* `contexte_general*.json`
* `*proper_names*.txt`
* `Participants.xls/.xlsx`
* `Sujets.xls/.xlsx`

11. **Génère `infos_projet.json`**

* localement ;
* et dans le dossier de transcription cible.

12. **Crée une copie de compatibilité audio**

* si un WAV mono 16 kHz existe, il le duplique en `audio_compatible.wav`.

### Finalité

Ce script sert donc à **normaliser et centraliser un lot de captation** (photos + audio + transcriptions + contexte) dans l’arborescence d’une affaire, en s’appuyant sur plusieurs scripts Python externes.

### Point d’attention

Il existe probablement une **erreur de variable** au début :

```bat
set "ROOT_PC=%~2"
if "%ROOT_DST%"=="" set "ROOT_DST=\\192.168.1.20\volume1\Affaires"
```

Le script lit `%~2` dans `ROOT_PC`, mais teste ensuite `ROOT_DST`, qui n’a pas été alimentée. En l’état, l’argument 2 risque de ne pas être utilisé comme prévu.

Je peux aussi vous faire une **lecture ligne par ligne** ou vous proposer une **version commentée et sécurisée** du script.
