# Les algorithmes dealflow Leonard / VINCI

Ce dossier contient les scripts Python pour structurer et qualifier un dealflow startups en trois étapes :

1. Algo 1 : préqualification initiale
2. Algo 2 : affinage stratégique VINCI
3. Algo 3 : orientation programme

Un script pipeline permet aussi d'enchaîner les trois algos startup par startup, sans attendre qu'un fichier complet soit terminé.

## Configuration

Les paramètres sont dans le fichier `.env`.

Variables principales :

```env
AZURE_OPENAI_DEPLOYMENT=gpt-5-mini-2
AZURE_OPENAI_API_VERSION=2024-12-01-preview
DEALFLOW_PROFILE=EUROPE
```

Les trois algos utilisent le même déploiement Azure OpenAI : `gpt-5-mini-2`.

### Profils Europe / LATAM

Le profil par défaut est `EUROPE` : il conserve les critères actuels.

Pour traiter un fichier LATAM :

```env
DEALFLOW_PROFILE=LATAM
```

En PowerShell, pour un lancement ponctuel :

```powershell
$env:DEALFLOW_PROFILE="LATAM"
python -B Pipeline_Algo123.py "mon_fichier_source.xlsx"
```

Le profil `LATAM` :

- demande des justifications en anglais ;
- exclut les startups principalement centrées sur construction of buildings et real estate ;
- inclut les autres sujets construction pertinents : infrastructure construction, civil engineering/public works, roads, bridges, tunnels, highways, rail/railway infrastructure et construction materials ;
- garde les traces dans `dealflow_profile` et `profile_exclusion_matches` ;
- utilise pour Algo3 les seuils ARR provisoires : Seed `< 500k EUR`, Catalyst `> 500k EUR`.

Dans l'app Streamlit, le profil se choisit directement dans la barre latérale.

Les scripts peuvent éviter des appels API quand l'Excel contient déjà assez d'informations :

```env
ALGO1_ENABLE_EXCEL_SHORTCUT=1
ALGO2_ENABLE_EXCEL_SHORTCUT=1
ALGO1_DISABLE_API=0
ALGO2_DISABLE_API=0
```

Si tu veux un mode ultra-économe pour Algo1 et Algo2 :

```env
ALGO1_DISABLE_API=1
ALGO2_DISABLE_API=1
```

Dans ce mode, si l'Excel ne suffit pas, la ligne part en vérification au lieu d'appeler Azure.

## App equipe interne

### `App_dealflow.py`

Interface web interne pour l'equipe. Elle permet de :

- uploader un fichier Excel source ;
- choisir le profil de traitement `Europe` ou `LATAM` ;
- lancer le pipeline Algo1 -> Algo2 -> Algo3 ;
- suivre les logs pendant l'analyse ;
- telecharger l'Excel final ;
- telecharger et visualiser le dashboard HTML ;
- retrouver les derniers runs.

Installation sur une machine interne :

```powershell
python -m pip install -r requirements.txt
```

Lancement simple sur Windows :

```powershell
.\Lancer_app_equipe.bat
```

Ou commande directe :

```powershell
python -m streamlit run App_dealflow.py --server.address 0.0.0.0 --server.port 8501
```

L'equipe ouvre ensuite :

```text
http://adresse-du-serveur:8501
```

Important : ne pas partager le vrai fichier `.env`. Utilise `.env.example`
comme modele, puis mets la vraie cle Azure OpenAI uniquement sur la machine
ou le serveur qui heberge l'app.

Deploiement recommande :

1. machine interne ou VM VINCI avec Python ;
2. dossier code synchronise ou depose par IT ;
3. `.env` serveur non partage ;
4. lancement via `Lancer_app_equipe.bat` ou service Windows ;
5. acces equipe via URL interne.

## Scripts principaux

### `Algo1.py`

Objectif : premier tri d'éligibilité.

Sorties :

- `A_GARDER`
- `A_VERIFIER`
- `A_ECARTER`

Commande :

```powershell
python -B Algo1.py "mon_fichier.xlsx"
```

Sortie :

```text
mon_fichier_algo_1_prequalification.xlsx
```

Sauvegarde temporaire :

```text
mon_fichier_algo_1_prequalification_sauvegarde_temp.xlsx
```

### `Algo2.py`

Objectif : filtre d'affinage stratégique après Algo1.

Par défaut, Algo2 traite seulement les startups `A_GARDER` de l'Algo1.

Sorties :

- `PREQUALIFIEE`
- `A_VERIFIER`
- `NON_PREQUALIFIEE`

Commande :

```powershell
python -B Algo2.py "mon_fichier_algo_1_prequalification.xlsx"
```

Recherche web Algo2 :

```env
ALGO2_TARGETED_RESEARCH=1
ALGO2_RESEARCH_MAX_PAGES=5
ALGO2_RESEARCH_MAX_LINKS_TO_TRY=15
ALGO2_RESEARCH_TEXT_PER_PAGE_CHARS=1400
```

Avec ce mode, Algo2 privilegie les pages utiles pour la decision :
clients/customers, case studies, use cases, industries served, product pages
et solutions. Les preuves trouvees sont ajoutees dans des colonnes dediees.

Sortie :

```text
mon_fichier_algo_2_affinage_strategique.xlsx
```

### `Algo3.py`

Objectif : orienter les startups préqualifiées vers le bon programme.

Par défaut, Algo3 traite seulement :

```env
ALGO3_ELIGIBLE_DECISIONS=PREQUALIFIEE
```

Orientations :

- `Seed`
- `Catalyst`
- `Matériaux`
- `Autre`
- `A_VERIFIER`

Commande :

```powershell
python -B Algo3.py "mon_fichier_algo_2_affinage_strategique.xlsx"
```

Recherche maturite Algo3 :

```env
ALGO3_TARGETED_MATURITY_RESEARCH=1
ALGO3_RESEARCH_MAX_PAGES=5
ALGO3_RESEARCH_MAX_LINKS_TO_TRY=15
ALGO3_RESEARCH_TEXT_PER_PAGE_CHARS=1400
```

Avec ce mode, Algo3 privilegie les pages utiles pour orienter Seed / Catalyst /
Materiaux : about, team, careers, customers, case studies, product, technology,
deployments, press, investors, manufacturing, production et materials. Les preuves
sont exportees dans des colonnes dediees, avec des scores locaux Seed, Catalyst,
Materiaux et maturite.

Sortie :

```text
mon_fichier_algo_3_orientation_programme.xlsx
```

## Pipeline automatique

### `Pipeline_Algo123.py`

Ce script enchaîne les trois algos startup par startup.

Logique :

```text
Startup -> Algo1
si A_GARDER -> Algo2
si PREQUALIFIEE -> Algo3
sinon arrêt avec trace
```

Commande :

```powershell
python -B Pipeline_Algo123.py "mon_fichier_source.xlsx"
```

Sortie :

```text
mon_fichier_source_pipeline_algo_1_2_3.xlsx
```

Sauvegarde temporaire :

```text
mon_fichier_source_pipeline_algo_1_2_3_sauvegarde_temp.xlsx
```

Le pipeline sauvegarde automatiquement toutes les 5 startups :

```env
PIPELINE_SAVE_EVERY=5
PIPELINE_RESUME_FROM_TEMP=1
PIPELINE_GENERATE_DASHBOARD=1
```

Si le script coupe, relance la même commande : il reprend depuis la sauvegarde.

Sheets principales du pipeline :

- `Synthese equipe`
- `Toutes les startups`
- `Stop Algo1 A verifier`
- `Stop Algo1 A ecarter`
- `Stop Algo2 A verifier`
- `Stop Algo2 Non prequal`
- `Algo1 A garder`
- `Algo2 Prequalifiees`
- `Algo3 Seed`
- `Algo3 Catalyst`
- `Algo3 Materiaux`
- `Algo3 Autre`
- `Algo3 A verifier`

`Synthese equipe` est la vue courte pour l'equipe : decision Algo1 et raison,
decision Algo2 et raison, orientation Algo3 et raison, statut pipeline,
prochaine action et principaux champs/scores utiles.

### `Generer_dashboard.py`

Ce script genere un dashboard HTML visuel a partir du fichier pipeline.
Le pipeline le genere automatiquement en fin de traitement si
`PIPELINE_GENERATE_DASHBOARD=1`.

Commande :

```powershell
python -B Generer_dashboard.py "mon_fichier_source_pipeline_algo_1_2_3.xlsx"
```

Sortie :

```text
mon_fichier_source_pipeline_algo_1_2_3_dashboard.html
```

Le dashboard contient les KPI, le funnel Algo1 -> Algo2 -> Algo3, les
repartitions par decision/orientation, les secteurs, les matrices d'enjeux,
les startups a revoir en priorite, les top opportunites et une table filtrable.
Il s'ouvre directement dans un navigateur.

## Optimisations coût / vitesse

Les algos utilisent plusieurs niveaux avant d'appeler Azure :

1. données Excel
2. mots-clés métier
3. score sémantique local
4. mémoire d'apprentissage
5. recherche web
6. API Azure OpenAI si nécessaire

Le score sémantique local est dans :

```text
semantic_prefilter.py
```

Il fonctionne sans dépendance lourde. Actuellement, le backend actif est `hashing`.

Un vrai modèle `sentence-transformers` pourra être activé plus tard si les dépendances sont installées :

```env
LOCAL_SEMANTIC_USE_TRANSFORMER=1
LOCAL_SEMANTIC_MODEL=sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2
```

## Apprentissage par corrections

Le système ne réentraîne pas GPT. Il apprend depuis tes corrections manuelles via une mémoire locale.

Fichiers :

- `learning_memory.py` : moteur de mémoire
- `learning_memory.json` : base d'exemples corrigés
- `Apprendre_feedback.py` : apprend depuis un Excel corrigé
- `Evaluer_et_apprendre.py` : évalue les erreurs et apprend

Dans le fichier pipeline, tu peux remplir :

- `manual_algo1_decision`
- `manual_algo2_decision`
- `manual_algo3_orientation`
- `manual_comment`

Valeurs attendues :

```text
manual_algo1_decision : A_GARDER / A_VERIFIER / A_ECARTER
manual_algo2_decision : PREQUALIFIEE / A_VERIFIER / NON_PREQUALIFIEE
manual_algo3_orientation : Seed / Catalyst / Matériaux / Autre / A_VERIFIER
```

Pour apprendre depuis un fichier corrigé :

```powershell
python -B Apprendre_feedback.py "mon_fichier_pipeline_corrige.xlsx"
```

Pour évaluer et apprendre en même temps :

```powershell
python -B Evaluer_et_apprendre.py "mon_fichier_pipeline_corrige.xlsx"
```

Sortie :

```text
mon_fichier_pipeline_corrige_rapport_evaluation_apprentissage.xlsx
```

Le rapport contient :

- `Synthese` : accuracy par algo
- `Apprentissage` : nombre d'exemples ajoutés
- `Distributions` : répartition des décisions
- `A corriger priorite` : lignes les plus utiles à relire
- `Erreurs Algo1`
- `Erreurs Algo2`
- `Erreurs Algo3`

## Mémoire d'apprentissage

Le fichier :

```text
learning_memory.json
```

stocke uniquement les corrections humaines validées.

Réglages :

```env
LEARNING_ENABLED=1
LEARNING_MEMORY_PATH=learning_memory.json
LEARNING_MATCH_THRESHOLD=0.28
LEARNING_MAX_EXAMPLES=5000
LEARNING_MAX_TEXT_CHARS=8000
```

Les algos ajoutent des traces :

- `algo1_learning_match_label`
- `algo1_learning_match_score`
- `algo1_learning_match_source`
- `algo2_learning_match_label`
- `algo2_learning_match_score`
- `algo2_learning_match_source`

## Fichiers du dossier

### Code

- `Algo1.py` : préqualification initiale
- `Algo2.py` : affinage stratégique
- `Algo3.py` : orientation programme
- `Pipeline_Algo123.py` : orchestration des trois algos
- `dealflow_profiles.py` : profils Europe / LATAM, exclusions LATAM et seuils ARR
- `semantic_prefilter.py` : scoring sémantique local
- `learning_memory.py` : mémoire d'apprentissage
- `Apprendre_feedback.py` : ingestion des corrections humaines
- `Evaluer_et_apprendre.py` : évaluation, rapport d'erreurs et apprentissage

### Configuration

- `.env` : variables Azure, limites de recherche, sauvegardes, apprentissage

### Données et sorties

- `*.xlsx` : fichiers Excel sources et résultats
- `*_sauvegarde_temp.xlsx` : sauvegardes de reprise
- `*_pipeline_algo_1_2_3.xlsx` : résultats pipeline
- `*_rapport_evaluation_apprentissage.xlsx` : rapports d'évaluation

### Dossiers techniques

- `__pycache__` : cache Python
- `.git`, `.codex`, `.agents` : dossiers techniques du projet

## Workflow conseillé

1. Lancer le pipeline sur un fichier source.

```powershell
python -B Pipeline_Algo123.py "mon_fichier_source.xlsx"
```

2. Ouvrir le résultat pipeline.
3. Corriger 50 à 200 lignes dans les colonnes `manual_*`.
4. Évaluer et apprendre.

```powershell
python -B Evaluer_et_apprendre.py "mon_fichier_source_pipeline_algo_1_2_3.xlsx"
```

5. Relancer le pipeline sur un nouveau fichier.

Les corrections précédentes seront utilisées pour réduire les erreurs et limiter les appels API.

## Conseils pour gros fichiers

Pour 2000 startups :

- garder `MAX_WORKERS=1`
- garder les sauvegardes auto
- éviter de lancer deux gros pipelines en même temps
- commencer par un test :

```env
PIPELINE_LIMIT_ROWS=100
```

Puis remettre :

```env
PIPELINE_LIMIT_ROWS=0
```

quand tout est stable.
