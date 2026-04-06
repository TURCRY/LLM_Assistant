Voici un **README.md court et opérationnel** pour le dossier `config/` de `LLM_Assistant`.

---

## `C:\LLM_Assistant\config\README.md`

```md
# Configuration — LLM_Assistant

## 1. Objet

Ce dossier contient les fichiers de configuration utilisés par `LLM_Assistant`.

Il distingue :

- les fichiers **locaux** au laptop ;
- les fichiers **appelés à être centralisés** côté serveur Flask (`GPT4All_Local`).

---

## 2. Principe d’architecture

`LLM_Assistant` est un **client**.

La logique cible est :

- le serveur Flask (`GPT4All_Local`) constitue la **source de vérité** ;
- les traitements structurants (création d’affaire, index, synchronisation, etc.) sont exécutés côté serveur ;
- le laptop prépare, déclenche et contrôle.

---

## 3. Statut des fichiers

### 3.1 Fichiers locaux (source de vérité locale)

Ces fichiers sont propres au client et doivent être versionnés :

- `config.json`
- `llm_scenarios.json`
- `prompt_tooltips.json`
- `system_prompt.json`
- `voxtral_report_prompts.json`

Ils définissent :

- le comportement de l’interface ;
- les scénarios LLM ;
- les prompts.

---

### 3.2 Fichier particulier — projets_index.json

```

config/projets_index.json

```

#### Statut

Ce fichier est :

- présent dans le dépôt ;
- utilisé localement ;
- **non canonique en production**.

#### Rôle

Il sert :

- de référence locale ;
- de fallback éventuel ;
- de support de développement.

#### Cible

La cible est que ce fichier soit lu côté serveur Flask, depuis :

```

\192.168.0.155\GPT4All_Local\config\projets_index.json

```

via la route `/create_affaire` (ou équivalent).

#### Règle

```

Le projets_index.json local ne doit pas être considéré comme la source de vérité.

```

---

## 4. Conséquences pour le développement

Toute évolution doit respecter :

- la délégation progressive des responsabilités vers le serveur ;
- l’absence de duplication des sources de vérité ;
- la compatibilité avec les futures routes Flask.

---

## 5. Règles Codex

Codex peut :

- modifier les fichiers de ce dossier ;
- améliorer leur structure ;
- préparer leur utilisation côté client.

Codex ne doit pas :

- considérer ces fichiers comme canoniques serveur ;
- introduire une dépendance forte à leur version locale ;
- dupliquer la logique serveur dans le client.

---

## 6. Principe directeur

```

Configuration locale = aide et pilotage
Configuration serveur = source de vérité

```
```

---

## Pourquoi ce README est important

Il permet d’éviter trois erreurs classiques :

* ❌ considérer le `projets_index.json` local comme officiel
* ❌ coder une logique client qui court-circuite le serveur
* ❌ dupliquer la logique métier entre laptop et PC fixe

---

## Résultat

Avec ce README + la note Codex précédente :

👉 tu verrouilles :

* la **source de vérité côté serveur**
* la **lecture seule côté laptop**
* la **préparation du futur `/create_affaire`**

---

