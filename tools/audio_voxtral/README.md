# Prétraitement audio Voxtral

Cette brique prépare les WAV natifs avant leur envoi au pipeline ASR :

- `prep_voxtral.bat` convertit chaque source en WAV mono PCM 16 bits, 16 kHz ;
- `assemblage.ps1` concatène plusieurs WAV préparés à partir d’un `liste.txt` généré dans le dossier audio de travail ;
- `assembler_wav.bat` conserve le lancement manuel historique et résout désormais le script PowerShell relativement à ce dossier.

Depuis le panneau **Voxtral (ASR / CR)**, renseigner un ou plusieurs chemins WAV, puis utiliser **Préparer audio Voxtral**. Un fichier unique produit `*_mono16_16000Hz.wav`; plusieurs fichiers produisent un `*_complet.wav`. Ce fichier final est ensuite utilisé comme entrée ASR.

Prérequis : `ffmpeg` accessible dans le `PATH` et Windows PowerShell 5.1.
