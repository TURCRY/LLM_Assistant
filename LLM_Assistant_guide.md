
# 🧠 Guide utilisateur — LLM_Assistant

Version : `app_final.py`  
Dernière mise à jour : juillet 2025

---

## 1. 🎯 Objectif

LLM_Assistant est une interface locale (Streamlit) permettant de :

- Interroger un modèle LLM hébergé localement via un serveur Flask (`gpt4all_flask.py`)
- Utiliser des prompts simples ou structurés
- Exploiter un RAG local (corpus par projet)
- Activer ou non une recherche Web assistée
- Gérer plusieurs projets via un index centralisé

---

## 2. 🗂 Structure des projets

Chaque projet est référencé dans :
```
LLM_Assistant/config/projets_index.json
```

### Exemple de structure :

```json
{
  "id_projet": "expertise_catnat_versailles",
  "nom": "Expertise CatNat Versailles",
  "commentaire": "Analyse sécheresse 2022",
  "rag_dossier_laptop": "C:/.../fichiers_rag",
  "rag_dossier_pcfixe": "C:/Dossier_rag/...",
  "chemin_config": "C:/.../projet_config.json",
  "model_name": "Mistral_7B",
  "date_creation": "2025-07-01"
}
```

---

## 3. 🧩 Fonctionnalités principales

### ✅ Sélection d’un projet
- Menu déroulant basé sur `projets_index.json`
- Charge automatiquement le fichier `projet_config.json`
- Affiche le modèle associé et les chemins RAG

### ✅ Choix du mode de génération :
- **Standard** → route `/annoter`
- **RAG** → route `/annoter_rag` (avec `rag_dossier_pcfixe`)
- **Web** → route `/annoter_web` (avec `use_web = true`)

### ✅ Construction du prompt
- 5 champs guidés :
  - Objectif
  - Contexte
  - Format attendu
  - Contraintes
  - Exemples
- + la question centrale

### ✅ Envoi au serveur Flask
- Envoie le `payload` enrichi avec tous les paramètres
- Détection automatique du modèle
- Envoi conditionné : WOL → ping → POST

---

## 4. 🔁 Données échangées

### 🔼 Envoi (`POST`)
- `prompt`, `system`, `temperature`, `top_p`, `top_k`, `repeat_penalty`
- `max_tokens`, `model_name`, `project_id`
- `rag_dossier_pcfixe` (si RAG)
- `use_web` (si Web)

### 🔽 Réponse
- `response`
- `model_name`
- `model_switched`
- `sources` (Web, si activé)

---

## 5. 🧪 Fichier `.env` requis

```
API_KEY=...
SERVER_IP=...
PORT=...
MAC_ADDRESS=...
```

---

## 6. 📁 Dossiers utilisés

| Élément | Emplacement par défaut |
|--------|-------------------------|
| Config globale | `LLM_Assistant/config/` |
| Prompts sauvegardés | `projets/<nom>/prompts_sauvegardés/` |
| Logs Q&A | `projets/<nom>/qa_logs/` |
| Corpus RAG (local) | `fichiers_rag/` dans le dossier projet |
| Corpus RAG (serveur) | `C:/Dossier_rag/<id_projet>` sur le PC fixe |

---

## 7. 📌 Dépendances clés

- `streamlit`
- `requests`
- `python-dotenv`
- `watchdog` *(si activation de détection auto des changements)*

---

## 8. 🚀 Lancement

```bash
streamlit run app_final.py
```

---

## 9. 📞 Assistance

Contacter l'opérateur projet via `nicolas.turcry@ntu-consult.com`  
LLM utilisé : `Mistral_7B (gguf)` via `llama-cpp-python` / `gpt4all`

---
