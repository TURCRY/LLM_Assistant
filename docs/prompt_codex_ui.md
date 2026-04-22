
---

### Prompt Codex

Analyse le fichier `annotation_interface_gpt.py` afin de vérifier et corriger la prise en compte du fichier `contexte_general_photos.json` dans le pipeline VLM.

Contexte :

* Le batch utilise désormais deux champs dédiés au VLM dans le JSON :

  * `vlm_system`
  * `vlm_user`

Exemple de structure :
{
  "mission": "...",
  "etat_avancement": "...",
  "system": "...",
  "user": "...",
  "vlm_system": "Tu décris uniquement ce qui est visible sur l’image, en français, de manière factuelle et prudente.",
  "vlm_user": "Pour cette mission, porte une attention prioritaire aux éléments d’ouvrage, aux désordres apparents, aux raccords, conduites, fissures, déformations, humidité et inachèvements. N’accorde pas d’importance centrale aux objets accessoires hors sujet."
}

* Ces champs doivent être prioritaires pour le prompting VLM.
* À défaut, fallback sur `mission`, `system`, `user`.

Objectifs :

1. Identifier précisément dans `annotation_interface_gpt.py` :

   * où est construit le prompt VLM (system + user),
   * où est chargé `contexte_general_photos.json`.

2. Vérifier si :

   * `vlm_system` et `vlm_user` sont lus,
   * sinon, démontrer qu’ils sont ignorés actuellement.

3. Proposer une correction minimale et ciblée :

   * intégrer `vlm_system` et `vlm_user` dans la construction du prompt VLM,
   * conserver la compatibilité avec les anciens JSON (fallback).

4. Contraintes métier à respecter :

   * le contexte texte ne doit jamais être pris comme description de l’image,
   * il doit seulement guider l’attention,
   * éviter toute dérive vers des objets hors sujet (jouets, mobilier, etc.),
   * conserver une séparation claire :

     * cadrage VLM (contexte)
     * consigne VLM (prompt)
     * contexte spécifique (transcription).

5. Fournir :

   * un `diff` strict et minimal de `annotation_interface_gpt.py`,
   * sans modifier la logique LLM,
   * sans modifier le serveur Flask.

6. Vérifier que :

   * le prompt final envoyé à `/vision/describe` est cohérent avec celui du batch,
   * l’ordre est respecté : image → texte.

Sortie attendue :

* analyse courte (points 1 à 2),
* diff exact,
* explication synthétique du comportement après correction.

Important :

* ne pas refactoriser massivement,
* ne pas déplacer de logique côté serveur,
* rester strictement aligné avec le fonctionnement actuel du batch VLM.

---
