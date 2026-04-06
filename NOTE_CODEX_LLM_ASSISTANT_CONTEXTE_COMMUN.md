# Note Codex — LLM_Assistant, contexte commun et source de vérité serveur

## 1. Objet

Cette note fixe le cadre d’intervention de Codex sur le dépôt `LLM_Assistant`.

L’objectif est d’adapter `LLM_Assistant` en tenant compte :

- d’un contexte commun de référence provenant de `GPT4All_Local` ;
- d’une séparation stricte entre :
  - le code client laptop ;
  - le code serveur Flask ;
  - les fichiers de configuration locaux ;
  - les fichiers de configuration canoniques côté serveur.

## 2. Périmètre autorisé

Codex peut modifier :

- le code de `LLM_Assistant` ;
- ses fichiers de configuration locaux ;
- sa documentation locale.

Codex ne doit pas modifier directement :

- le code serveur `GPT4All_Local` ;
- les routes Flask du serveur ;
- les fichiers canoniques côté serveur.

En revanche, Codex peut :

- lire le contexte de référence `GPT4All_Local_Reference` ;
- proposer des adaptations minimales du serveur sous forme de note ou de patch proposé ;
- préparer `LLM_Assistant` à fonctionner avec les futures évolutions du serveur.

## 3. Règle d’architecture

Le laptop n’est pas la source de vérité du serveur.

Le principe cible est :

- `LLM_Assistant` agit comme client ;
- `GPT4All_Local` côté serveur lit la configuration canonique ;
- les créations d’affaire et traitements structurants doivent être déportés vers le serveur Flask via des routes dédiées.

## 4. Cas particulier de `projets_index.json`

### 4.1 Dans `LLM_Assistant`

Le fichier local :

`C:\LLM_Assistant\config\projets_index.json`

doit être considéré comme :

- un fichier de référence ;
- un fallback local éventuel ;
- un artefact utile au développement ou à la transition.

Il ne doit pas être considéré comme la source de vérité de production.

### 4.2 Côté serveur

La cible est que le serveur Flask lise le fichier canonique situé dans :

`\\192.168.0.155\GPT4All_Local\config\projets_index.json`

via la route `/create_affaire` ou d’autres routes serveur à créer.

Cette partie serveur sera gérée dans un autre dépôt.

## 5. Conséquence pour `LLM_Assistant`

Codex doit préparer `LLM_Assistant` pour que :

- la logique métier structurante soit déléguée au serveur ;
- les appels futurs à `/create_affaire` soient possibles ;
- la présence de `config/projets_index.json` local ne crée pas d’ambiguïté sur la source de vérité.

## 6. Règles de travail

- privilégier des modifications minimales et localisées ;
- ne pas casser les comportements existants ;
- ne pas supposer que le fichier local `projets_index.json` restera la source principale ;
- documenter explicitement toute hypothèse transitoire ;
- signaler clairement les points qui dépendent d’une future évolution du serveur.

## 7. Attendus concrets

Codex peut notamment :

- repérer où `LLM_Assistant` lit ou pourrait lire `projets_index.json` ;
- distinguer usage local / usage canonique distant ;
- préparer les appels HTTP vers le serveur ;
- ajouter des wrappers, helpers ou documentation de transition ;
- produire, si nécessaire, une proposition de patch pour `GPT4All_Local`, sans l’appliquer.

## 8. Principe directeur

`LLM_Assistant` doit rester compatible avec un serveur Flask centralisé, qui constitue la source de vérité de production.