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

### Profils Europe, LATAM et VINCI Concessions

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

#### VINCI Concessions

Dans la barre latérale de Streamlit, choisis **VINCI Concessions** dans **Profil d'analyse**, puis sélectionne une **Business unit**. Chaque analyse utilise la grille de la BU choisie. Le périmètre reprend les cinq divisions présentées par VINCI Concessions. [Rapport d'engagement VINCI Concessions](https://engagementreport.vinci-concessions.com/2024-2025/doc/article/6/)

| Business unit | Code `DEALFLOW_PROFILE` | Cas d'usage à repérer | Indicateurs possibles |
| --- | --- | --- | --- |
| **VINCI Airports** | `VINCI_AIRPORTS` | Parcours passager, escales et bagages, maintenance des infrastructures, environnement, commerces et accès. [Détail de la grille](#vinci-airports) | Attente, ponctualité, disponibilité, consommation |
| **VINCI Autoroutes** | `VINCI_AUTOROUTES` | Réseau concédé en France. Sécurité des patrouilleurs, détection d'incidents, inspection des chaussées et viaducs, trafic et télépéage, recharge et covoiturage sur les aires. [Activités du réseau](https://www.vinci-autoroutes.com/fr/actualites/) | Délai d'intervention, état des chaussées, disponibilité des bornes |
| **VINCI Highways** | `VINCI_HIGHWAYS` | Concessions routières internationales. Péage en flux libre, services de paiement, supervision du trafic, maintenance des ponts et tunnels, résilience du réseau. [Péage](https://vinci-highways.com/tolling-services/) · [Exploitation et maintenance](https://vinci-highways.com/operations-maintenance/) | Transactions correctement traitées, temps d'intervention, disponibilité des équipements |
| **VINCI Railways** | `VINCI_RAILWAYS` | Infrastructures ferroviaires et de transport guidé. Géométrie des voies, maintenance des aiguillages et caténaires, planification des interventions, accessibilité en gare. [Périmètre ferroviaire](https://engagementreport.vinci-concessions.com/2025-2026/fr/doc/article/13/) · [Expertise MESEA](https://www.mesea.fr/expertise/) | Défauts détectés, retards, durée des interruptions |
| **VINCI Stadium** | `VINCI_STADIUM` | Grandes enceintes sportives et événementielles. Flux et accessibilité, billetterie professionnelle, gestion des prestataires, entretien des installations, énergie, eau et déchets. [Périmètre des stades](https://engagementreport.vinci-concessions.com/2025-2026/fr/doc/article/13/) · [Engagements de l'Allianz Riviera](https://www.allianz-riviera.fr/fr/allianz-riviera/des-engagements-de-developpement-durable-forts) | Temps d'accès, disponibilité des équipements, consommation pendant les événements |

En ligne de commande, renseigne le code d'une BU dans `DEALFLOW_PROFILE`. Par exemple, pour VINCI Highways.

```powershell
$env:DEALFLOW_PROFILE="VINCI_HIGHWAYS"
python -B Pipeline_Algo123.py "mon_fichier_source.xlsx"
```

Les grilles constituent un premier ensemble de critères fondés sur les activités publiques des BU.

- **À garder** si un produit identifiable apporte une réponse concrète à un besoin de l'exploitant ou d'un partenaire professionnel de la BU. Une référence dans un domaine voisin peut suffire si le transfert est crédible.
- **À vérifier** si les données sont insuffisantes, contradictoires ou si le cas d'usage pour la BU reste à confirmer. La justification indique l'information à compléter.
- **À écarter** si l'activité décrite est clairement hors du périmètre de la BU. Par exemple, une application de scores sportifs destinée aux particuliers est hors périmètre pour VINCI Stadium.

Le raccourci Excel traite les cas explicites en français ou en anglais. Il repère un produit, son client professionnel et un usage de la grille choisie. Il peut aussi écarter certains cas grand public explicites. Les autres lignes passent à l'analyse approfondie avec Azure, ou à **À vérifier** quand `ALGO1_DISABLE_API=1`. Pour tester Algo1 sans Azure, utilise cet exemple après avoir sélectionné le profil dans PowerShell.

```powershell
$env:ALGO1_DISABLE_API="1"
python -B Algo1.py "mon_fichier_source.xlsx"
Remove-Item Env:ALGO1_DISABLE_API
```

Dans l'analyse approfondie, Algo1 vérifie que l'extrait justificatif figure dans l'Excel ou dans le texte du site. Dans la feuille **Synthese equipe**, `dealflow_profile` identifie la BU. Les colonnes `algo1_bu_metier`, `algo1_bu_cas_usage` et `algo1_bu_element_observe` retracent la décision. Les exports VINCI Airports utilisent le préfixe historique `algo1_airports_`. Les informations à compléter figurent dans `algo1_infos_manquantes` pour le pipeline, ou `missing_information` pour Algo1 seul.

Les lignes retenues poursuivent l'affinage stratégique et l'orientation selon les règles du pipeline Leonard. En cas de reprise, Streamlit réutilise la BU du lancement et le pipeline vérifie le profil de la sauvegarde. Les grilles sont définies dans `vinci_concessions.py` et `vinci_airports.py`.

#### VINCI Airports

Un outil de gestion des files en gare peut être retenu si son application à un terminal et le bénéfice pour l'exploitant sont étayés.

| Métier | Exemples de cas d'usage | Indicateurs à examiner |
| --- | --- | --- |
| Parcours passager et accessibilité | Prévision des files, réservation de créneaux de passage, orientation en terminal, assistance aux personnes à mobilité réduite. [Qualité de service](https://vinci-airports.com/en/our-expertises/quality-of-service/) | Temps d'attente, délai de prise en charge |
| Opérations au sol et bagages | Coordination des escales, suivi des bagages, gestion des équipements au sol. [Essentials 2025](https://essentials.vinci-airports.com/2025/en/doc/article/23/) | Durée d'escale, ponctualité, délai de livraison des bagages |
| Infrastructure et maintenance | Inspection des pistes, suivi des équipements, maintenance prédictive. [Modernisation des infrastructures](https://essentials.vinci-airports.com/2025/en/doc/article/20/) | Disponibilité des équipements, durée d'intervention |
| Énergie et environnement | Pilotage énergétique, recharge des véhicules de service, détection des fuites d'eau, valorisation des déchets, suivi de la biodiversité. [Engagements environnementaux](https://vinci-airports.com/en/our-commitments/) | Énergie, émissions, consommation d'eau, déchets évités |
| Commerces, parkings et accès | Prévision des flux commerciaux, gestion du taux d'occupation des parkings, réservation ou optimisation des navettes. [Activités extra-aéronautiques](https://vinci-airports.com/en/our-expertises/extra-aeronautical-activities/) | Taux d'occupation, temps d'accès, revenu par passager |

### Raccourcis Excel et appels API

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
- choisir `Europe`, `LATAM` ou une BU de `VINCI Concessions` ;
- lancer le pipeline Algo1 -> Algo2 -> Algo3 en arriere-plan ;
- suivre les logs pendant l'analyse ;
- telecharger l'Excel final ;
- telecharger et visualiser le dashboard HTML ;
- retrouver les derniers runs.

L'interface reprend l'identite Leonard / VINCI : logo officiel, couleurs
`#004489`, `#00B4FF` et `#FF005A`, titres en `Vinci Serif` et interface en
`Vinci Sans`. Sa structure visuelle s'inspire de la Constellation des tendances
2026 : fond orbital discret, parcours Algo1 -> Algo2 -> Algo3 et espace de
pilotage en deux zones. Les ressources locales necessaires sont conservees dans
`assets/` afin que le rendu reste stable sur Streamlit Cloud.

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

Depuis l'app, fermer l'onglet navigateur ne coupe pas le job en cours :
le pipeline continue en arriere-plan, ecrit `pipeline.log` et `run_status.json`
dans le dossier du run, puis l'app peut reprendre le suivi via "Analyses en cours".
Le bouton `Arreter l'analyse` demande l'arret du job et conserve la derniere
sauvegarde temporaire disponible.
Pour continuer, ouvre `Analyses interrompues`, puis clique sur
`Reprendre ce run`. Le pipeline repart apres la derniere ligne presente dans
la sauvegarde temporaire, avec le même profil que celui du lancement.
En revanche, si la machine ou le serveur qui heberge Streamlit s'eteint, le job
s'arrete aussi.

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

Les fichiers finaux disposent aussi d'une mise en forme de lecture :

- onglet `Mode d'emploi` avec la légende et l'action attendue ;
- vert pour les résultats validés, orange pour les vérifications et rouge pour
  les rejets ;
- couleurs distinctes pour Seed, Catalyst et Matériaux ;
- filtres, volets figés, liens cliquables et colonnes adaptées au contenu ;
- échelle colorée sur les scores de confiance.

La mise en forme complète n'est pas appliquée aux sauvegardes temporaires afin
de conserver des sauvegardes rapides pendant l'analyse.

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
