# LLM_Assistant
# Note de passation
## 2 juillet 2026

Auteur : Nicolas TURCRY / ChatGPT

---

# 1. Objet du document

Cette note a pour objet de décrire l'état réel du projet **LLM_Assistant** au 2 juillet 2026, les décisions d'architecture déjà prises, les développements validés, les incidents rencontrés ainsi que la stratégie de reprise des développements.

Ce document constitue le point d'entrée de tout nouveau développement sur le projet.

Il doit être lu avant toute modification importante de `app.py`.

---

# 2. Philosophie générale du projet

LLM_Assistant n'est plus un simple client Streamlit.

Le projet devient progressivement un **orchestrateur métier** destiné à assister un expert judiciaire tout au long du cycle de vie d'une expertise.

Le logiciel doit permettre notamment :

- création d'une affaire ;
- gestion des parties ;
- réception des documents des parties ;
- pré-traitement documentaire ;
- OCR ;
- renommage des pièces ;
- classement automatique ;
- captations audio ;
- captations photographiques ;
- transcription ASR ;
- annotation photographique ;
- génération de comptes-rendus LLM ;
- synchronisation NAS / Laptop / PC fixe.

Le logiciel doit masquer autant que possible la complexité technique au profit d'un workflow métier cohérent.

---

# 3. Architecture générale

Le projet repose sur trois machines.

## Laptop

Le laptop constitue le poste principal de l'expert.

Il permet :

- la gestion des affaires ;
- la préparation des traitements ;
- la consultation des résultats ;
- le pilotage des traitements lourds.

Le laptop ne doit pas réaliser lui-même les traitements IA coûteux.

---

## NAS

Le NAS constitue la référence documentaire.

Toutes les affaires y sont stockées.

Le NAS contient également les captations et les documents de travail.

---

## PC fixe

Le PC fixe dispose des ressources GPU.

Il exécute notamment :

- OCR
- ASR Voxtral
- AnnotationPhotosGPT
- traitements LLM lourds

Le PC fixe est considéré comme un serveur de calcul.

---

# 4. Architecture cible

Le laptop prépare les traitements.

↓

Le laptop dépose un job.

↓

Le PC fixe exécute le job.

↓

Les résultats sont écrits sur le NAS.

↓

Le laptop récupère les résultats.

Le laptop devient ainsi un poste de pilotage.

---

# 5. Organisation des affaires

Chaque affaire possède une structure canonique.

Les informations métier sont réparties principalement entre :

```
_Config
_DB
AA_Expert_Admin
AF_Expert_ASR
AE_Expert_captations
```

L'arborescence canonique est définie dans :

```
config/arborescence_canonique.yaml
```

---

# 6. Base SQLite

Une base SQLite est prévue par affaire.

Elle constitue progressivement la source métier principale.

Les journaux JSONL restent utilisés uniquement :

- pour compatibilité ;
- comme journaux techniques.

La logique métier doit progressivement migrer vers SQLite.

---

# 7. Workflow documentaire

Le workflow retenu est le suivant.

```
Réception d'un dépôt documentaire

↓

Création d'une cohorte

↓

OCR

↓

Analyse Dire/Bordereau

↓

Suggestions de renommage

↓

Validation par l'expert

↓

Attribution des codes pièces expert

↓

Écriture SQLite

↓

États 1 à 4
```

Cette séquence ne doit pas être inversée.

---

# 8. Cohorte documentaire

Une cohorte représente une transmission documentaire.

Une cohorte est définie par :

- une date de transmission ;
- un déposant ;
- un avocat ;
- une liste de documents.

Toutes les analyses ultérieures doivent porter par défaut sur cette cohorte uniquement.

---

# 9. Numérotation des pièces

Deux numérotations coexistent.

## Numéro avocat

Exemple :

```
Pièce n° 12
```

Il correspond au bordereau produit par l'avocat.

---

## Numéro expert

Format :

```
NN-NNNN
```

avec

```
NN
```

code du déposant

et

```
NNNN
```

compteur propre au déposant.

Le numéro expert est attribué uniquement lors de la validation du dépôt documentaire.

Jamais avant.

---

# 10. Synchronisation

La doctrine retenue est la suivante.

Le NAS constitue la référence documentaire.

Le PC fixe constitue un miroir de travail.

Le laptop ne conserve que les éléments utiles.

Les traitements intermédiaires volumineux ne doivent pas être synchronisés vers le laptop.

---

# 11. Spooler PC fixe

Une première version est opérationnelle.

Organisation :

```
queued
running
done
failed
logs
```

Le laptop fabrique des jobs JSON.

Le PC fixe les exécute.

Cette architecture doit être conservée.

---

# 12. Incident majeur

Pendant les développements de juin 2026, un incident PowerShell a vidé `app.py`.

Une restauration a été effectuée.

Conséquences :

- certaines fonctionnalités récentes ont disparu ;
- plusieurs blocs sont revenus à une version plus ancienne.

Le dépôt GitHub n'est pas responsable de cette perte.

---

# 13. Audit Git

Le dépôt Git est sain.

En revanche, plusieurs semaines de développements avaient été réalisées sans commits intermédiaires.

Ces développements n'ont donc pas pu être restaurés automatiquement.

Une discipline Git beaucoup plus stricte est désormais retenue.

---

# 14. Checkpoint Codex

Un checkpoint Codex a été retrouvé.

Après récupération, il compile.

Cependant il ne contient pas les développements les plus récents :

- DeepSeek OCR ;
- Spooler complet ;
- Débrief ;
- AnnotationPhotosGPT ;
- Cohortes documentaires ;
- SQLite avancée.

Il est conservé uniquement comme archive technique.

---

# 15. État actuel

Le fichier `app.py` actuel constitue la base de travail.

Il contient plusieurs développements récents mais présente encore des incohérences.

La stratégie retenue consiste à stabiliser progressivement cette version.

---

# 16. Priorités de reprise

Les développements seront repris dans l'ordre suivant.

1. Sélecteur d'affaires
2. Juridiction
3. Pré-traitement documentaire
4. Cohortes
5. SQLite
6. États 1 à 4
7. DeepSeek OCR
8. Captations
9. Voxtral
10. Débrief
11. AnnotationPhotosGPT

Chaque étape devra être validée avant de poursuivre.

---

# 17. Règles de développement

Les règles suivantes deviennent obligatoires.

## Une fonctionnalité = un commit

Ne jamais accumuler plusieurs jours de développement sans commit.

---

## Toujours compiler

Après chaque modification importante :

```
python -m py_compile app.py
```

---

## Toujours tester

Chaque fonctionnalité doit faire l'objet :

- d'un test unitaire ;
- d'un test fonctionnel.

---

## Toujours sauvegarder

Avant toute modification importante :

```
app.py
```

doit être sauvegardé.

---

## Une seule fonctionnalité à la fois

Ne jamais modifier simultanément plusieurs workflows métier.

---

# 18. Documents de référence

Les documents suivants doivent être conservés.

```
audit_app_after_restore.patch

diff_app_actuel.patch

app_from_codex_checkpoint_recovered.py

app_from_codex_checkpoint_recovered_report.json
```

Ils constituent les archives techniques de l'incident.

---

# 19. Conclusion

Le projet est désormais suffisamment mûr pour être considéré comme un véritable système d'information métier.

L'objectif des prochaines semaines n'est plus d'ajouter rapidement des fonctionnalités, mais de :

- consolider l'architecture ;
- stabiliser les workflows ;
- documenter les règles métier ;
- renforcer la stratégie Git.

La qualité et la traçabilité des développements deviennent désormais prioritaires par rapport à la vitesse de développement.