# PPTX Service

Micro-service qui génère et modifie des présentations PowerPoint via LLM.
Conçu pour s'intégrer à SiaGPT (via MCP ou API REST).

**Principe clé** : le LLM ne génère jamais de code — il lit et retourne du XML PowerPoint directement. Zéro `exec()`, zéro risque d'exécution arbitraire.

---

## Comment ça marche (vue d'ensemble)

### Petit rappel : un fichier .pptx, c'est quoi ?

Un fichier PowerPoint `.pptx` n'est rien d'autre qu'un **fichier ZIP** contenant des fichiers XML. Si tu renommes `presentation.pptx` en `presentation.zip` et que tu l'ouvres, tu verras :

```
presentation.zip/
├── [Content_Types].xml          ← "Registre" : liste tous les fichiers et leur type
├── _rels/.rels                  ← Liens entre fichiers (qui référence qui)
├── ppt/
│   ├── presentation.xml         ← La "table des matières" (ordre des slides)
│   ├── slides/
│   │   ├── slide1.xml           ← Le contenu de chaque slide (texte, positions, styles)
│   │   ├── slide2.xml
│   │   └── ...
│   ├── slides/_rels/
│   │   ├── slide1.xml.rels      ← Les liens de la slide 1 (layout, images, notes)
│   │   └── ...
│   ├── slideLayouts/            ← Les modèles de mise en page
│   ├── slideMasters/            ← Le style global (couleurs, polices du thème)
│   ├── theme/                   ← La palette de couleurs et polices
│   └── media/                   ← Les images embarquées
└── docProps/                    ← Métadonnées (auteur, date, etc.)
```

Notre service travaille directement sur ces fichiers XML — c'est comme ça qu'on modifie le texte, les styles et la structure sans jamais casser le formatage.

### Les deux LLM

Il y a deux LLM dans le système, avec des rôles distincts :

- **Le Chef** = le LLM de SiaGPT (celui à qui l'utilisateur parle dans le chat). Il comprend la demande, choisit le bon template, décide d'appeler `generate_pptx` ou `edit_pptx`. Il ne touche jamais au PPTX lui-même.
- **L'Ouvrier** = le LLM appelé par ce service (via l'API `/chat/plain_llm`). Il reçoit du XML brut et des instructions techniques, et retourne du XML modifié. Il ne sait rien de la collection, des templates, ni de l'utilisateur.

### Où sont stockés les templates Sia ?

Les templates Sia Partners sont des fichiers `.pptx` stockés dans **SiaGPT Medias** (même système que les fichiers utilisateur). Chaque template a un UUID.

```
SiaGPT Medias (collection)
├── 📄 abc-111-...  Template Sia - Proposition commerciale.pptx
├── 📄 abc-222-...  Template Sia - Comité de pilotage.pptx
├── 📄 abc-333-...  Template Sia - Rapport de mission.pptx
├── 📄 xyz-444-...  ma-presentation-modifiee.pptx  (fichier utilisateur)
└── ...
```

**C'est le Chef qui connaît les templates** (via son system prompt). Quand l'utilisateur dit "fais-moi une propale", le Chef sait qu'il faut utiliser le template "Proposition commerciale" et passe son UUID au service.

### Le workflow complet

#### Diagramme visuel (rendu par GitHub)

```mermaid
sequenceDiagram
    participant U as 👤 Utilisateur
    participant S as 🧠 SiaGPT<br/>(Le Chef)
    participant M as 📦 SiaGPT<br/>Medias
    participant P as ⚙️ PPTX<br/>Service
    participant L as 🤖 LLM<br/>Ouvrier

    U->>S: "Fais-moi une propale pour Airbus"

    Note over S: Le Chef connaît les templates.<br/>Il choisit "Proposition commerciale"<br/>UUID = abc-111-...

    alt Création avec template
        S->>P: generate_pptx(prompt, template_file_id)
        P->>M: GET /medias/{template_file_id}/download
        M-->>P: template.pptx
    else Création sans template
        S->>P: generate_pptx(prompt)
        Note over P: Crée un squelette vierge
    else Édition d'un fichier existant
        S->>P: edit_pptx(prompt, source_file_id)
        P->>M: GET /medias/{source_file_id}/download
        M-->>P: fichier.pptx
    end

    Note over P: 1. UNPACK → XML

    P->>L: Structure + prompt (Phase 1)
    L-->>P: Plan JSON

    loop Chaque slide du plan
        P->>L: XML slide + instructions (Phase 2)
        L-->>P: XML modifié
    end

    Note over P: CLEAN → VALIDATE → PACK

    P->>M: POST /medias/ (pptx + collection_id)
    M-->>P: {uuid: "xyz-999-..."}

    P-->>S: {status: ok, media_uuid: "xyz-999-..."}
    S-->>U: "Voilà ta propale ! 📎"
```

#### Tous les inputs/outputs du service

```
INPUTS (ce que le Chef envoie au service)
─────────────────────────────────────────
┌─────────────────────────────────────────────────────────────────────┐
│  generate_pptx                                                      │
│  ├── prompt            (requis)  "Crée une propale pour Airbus"     │
│  └── template_file_id  (option)  "abc-111-..." UUID du template     │
│                                  Si omis → squelette vierge         │
├─────────────────────────────────────────────────────────────────────┤
│  edit_pptx                                                          │
│  ├── prompt            (requis)  "Change les couleurs en bleu"      │
│  └── source_file_id    (requis)  "xyz-444-..." UUID du fichier      │
└─────────────────────────────────────────────────────────────────────┘

VARIABLES D'ENVIRONNEMENT (configurées au déploiement)
──────────────────────────────────────────────────────
┌─────────────────────────────────────────────────────────────────────┐
│  LLM_API_KEY           Bearer token pour appeler /chat/plain_llm    │
│  LLM_API_URL           https://backend.siagpt.ai/chat/plain_llm    │
│  LLM_MODEL             claude-4.5-sonnet                            │
│  SIAGPT_MEDIAS_URL     https://backend.siagpt.ai/medias             │
│  SIAGPT_COLLECTION_ID  UUID de la collection cible pour les uploads │
│  MAX_RETRIES           4 (tentatives si XML invalide)               │
└─────────────────────────────────────────────────────────────────────┘

OUTPUT (ce que le service retourne au Chef)
───────────────────────────────────────────
{
  "status": "ok",
  "media_uuid": "xyz-999-...",        ← UUID du fichier créé/modifié
  "media_name": "propale_airbus.pptx",
  "summary": "Création de 8 slides pour proposition commerciale Airbus",
  "modified_slides": ["slide1.xml", "slide2.xml", ...],
  "added_slides": ["slide6.xml", "slide7.xml"],
  "removed_slides": ["slide5.xml"],
  "errors": []                        ← vide si tout va bien
}
```

#### Le parcours du fichier PPTX (étape par étape)

```mermaid
graph TD
    A[📦 SiaGPT Medias<br/>template.pptx] -->|"GET /medias/{uuid}/download"| B["1️⃣ UNPACK<br/>ZIP → dossier XML<br/>+ pretty-print<br/>+ escape smart quotes"]
    B --> C["2️⃣ INSPECT<br/>Lire structure :<br/>slides, shapes, textes,<br/>positions, layouts"]
    C --> D["3️⃣ PLANIFIER<br/>🤖 LLM Ouvrier Phase 1<br/><br/>Input : structure JSON + prompt<br/>Output : plan JSON"]
    D --> E["4️⃣ MODIFIER<br/>🤖 LLM Ouvrier Phase 2<br/><br/>Pour chaque slide :<br/>Input : XML + instructions<br/>Output : XML modifié<br/>⟲ Retry si invalide (max 4x)"]
    E --> F["5️⃣ CLEAN<br/>Supprimer orphelins<br/>MAJ Content_Types"]
    F --> G["6️⃣ VALIDATE<br/>8 checks structurels<br/>+ validation XSD<br/>+ auto-repair"]
    G --> H["7️⃣ PACK<br/>Condenser XML<br/>Restaurer smart quotes<br/>→ fichier .pptx"]
    H -->|"POST /medias/<br/>+ collection_id"| I["📦 SiaGPT Medias<br/>résultat.pptx<br/>UUID = xyz-999-..."]

    style A fill:#4a90d9,color:#fff
    style I fill:#27ae60,color:#fff
    style D fill:#f39c12,color:#fff
    style E fill:#f39c12,color:#fff
```

---

## Les outils PPTX en détail

### pptx_tools.py — Manipulation des fichiers

Ce module sait ouvrir, fermer et manipuler les fichiers PPTX. Il ne sait rien du LLM — c'est de la plomberie pure.

#### `unpack(pptx_bytes, output_dir) → str`

**Ce que ça fait** : décompresse le fichier .pptx (qui est un ZIP) dans un dossier, et rend le XML lisible.

**Pourquoi** : le XML brut de PowerPoint est minifié (tout sur une ligne, illisible). L'unpack le met en forme pour que le LLM puisse le lire et le modifier correctement.

**En plus** : escape les "smart quotes" (`"` `"` `'` `'`) en entités XML (`&#x201C;` etc.) pour éviter les problèmes d'encodage quand le LLM modifie le texte.

```
presentation.pptx (ZIP binaire)
        │
        ▼  unpack()
/tmp/unpacked/
├── [Content_Types].xml  ← XML proprement indenté
├── ppt/slides/slide1.xml  ← Lisible par le LLM
└── ...
```

#### `pack(unpacked_dir, original_bytes) → bytes`

**Ce que ça fait** : l'opération inverse de unpack — repackage le dossier en fichier .pptx.

**Pourquoi c'est pas juste un zip** : avant de zipper, il faut :
1. **Condenser le XML** : retirer l'indentation qu'on a ajoutée (PowerPoint peut mal gérer les espaces parasites)
2. **Restaurer les smart quotes** : remettre les vrais caractères Unicode
3. **Préserver la compression** : si on a le fichier original, on réutilise ses niveaux de compression pour chaque fichier interne (sinon PowerPoint peut se plaindre)

#### `clean(unpacked_dir) → list[str]`

**Ce que ça fait** : le grand ménage avant de repackager. Supprime tout ce qui ne devrait plus être là.

**Les 5 nettoyages** :
1. **Slides orphelines** : slides qui existent dans `ppt/slides/` mais ne sont plus référencées dans `presentation.xml` (ex : on a supprimé une slide du plan mais le fichier XML traîne encore)
2. **Fichiers .rels orphelins** : fichiers de relations qui n'ont plus de fichier parent
3. **Dossier poubelle** : PowerPoint crée parfois un dossier `Trash/` — on le supprime
4. **Fichiers non-référencés** : images, médias, notes qui ne sont référencés par aucun .rels
5. **Mise à jour Content_Types** : après suppression de fichiers, met à jour le registre `[Content_Types].xml`

**Pourquoi c'est critique** : sans ce nettoyage, PowerPoint affiche le message "Ce fichier est endommagé — voulez-vous le réparer ?" et peut perdre du contenu.

#### `duplicate_slide(unpacked_dir, source_filename) → dict`

**Ce que ça fait** : crée une copie exacte d'une slide existante, avec tout ce qui va avec.

**Pourquoi c'est compliqué** : dupliquer une slide dans un PPTX, ce n'est pas juste copier un fichier. Il faut :
1. Copier le XML de la slide (`slide3.xml` → `slide4.xml`)
2. Copier son fichier de relations (`.rels`)
3. Copier ses notes (si elle en a)
4. Générer de nouveaux IDs uniques (slide ID, relationship ID)
5. Enregistrer le nouveau fichier dans `[Content_Types].xml`
6. (Optionnel) l'ajouter dans `presentation.xml` à la bonne position

Retourne un dict avec les IDs générés pour pouvoir l'insérer dans la présentation.

#### `add_slide_to_presentation(unpacked_dir, sld_id, r_id, position) → None`

**Ce que ça fait** : insère une slide dans l'ordre de la présentation en modifiant `presentation.xml` et son `.rels`.

**Contexte** : `duplicate_slide` crée les fichiers mais ne touche pas à l'ordre. Cette fonction s'en charge — elle ajoute l'entrée `<p:sldId>` dans `<p:sldIdLst>` à la position voulue.

---

### pptx_validate.py — Validation complète

Ce module vérifie que le PPTX n'est pas corrompu après modification. Deux niveaux.

#### Niveau 1 — Checks structurels

| Check | Ce qu'il vérifie | Exemple d'erreur détectée |
|-------|------------------|--------------------------|
| **XML bien formé** | Chaque fichier XML se parse sans erreur | Tag non fermé, caractère invalide |
| **Namespaces** | Les préfixes dans `mc:Ignorable` sont déclarés | LLM qui retire un namespace du root element |
| **IDs uniques** | Pas de doublons dans les IDs de shapes et slides | Deux shapes avec `id="5"` dans la même slide |
| **Références .rels** | Chaque lien pointe vers un fichier existant | `.rels` qui pointe vers `slide999.xml` inexistant |
| **Content_Types** | Tous les fichiers importants sont déclarés | Slide ajoutée mais pas dans `[Content_Types].xml` |
| **Slide layouts** | Chaque layout référencé existe dans les relations | `r:id` qui ne correspond à rien |
| **Pas de doublons** | 1 seul slideLayout par slide | Bug de duplication qui crée 2 layouts |
| **Notes non partagées** | 1 notesSlide par slide maximum | 2 slides qui pointent vers la même note |

#### Niveau 2 — Validation XSD

**XSD = XML Schema Definition.** Ce sont les schémas officiels de Microsoft qui définissent la "grammaire" du format PPTX. Par exemple, le schema `pml.xsd` dit : "un `<p:sld>` peut contenir un `<p:cSld>`, qui peut contenir un `<p:spTree>`, etc."

Si le LLM invente un tag (`<p:monTrucInventé>`), les checks structurels ne le voient pas (c'est du XML valide). Mais la validation XSD le détecte immédiatement.

**Comparaison avec l'original** : les templates ont souvent des erreurs XSD pré-existantes (extensions Microsoft non-standard). Notre validateur compare avec le fichier original et ne remonte que les **nouvelles** erreurs introduites par nos modifications.

#### Auto-repair

`xml:space="preserve"` : si un texte commence ou finit par un espace (`" Texte"`, `"Texte "`), PowerPoint le supprime silencieusement à l'ouverture sauf si `xml:space="preserve"` est présent sur le tag `<a:t>`. Notre validateur l'ajoute automatiquement.

---

## Structure du projet

```
pptx-service/
├── main.py                ← Service FastAPI : REST + MCP + orchestration workflow
├── pptx_tools.py          ← Manipulation PPTX : unpack, pack, clean, duplicate
├── pptx_validate.py       ← Validation : structurelle + XSD
├── schemas/               ← Schemas XSD Office Open XML (dans Docker)
├── system_prompt.md       ← Instructions pour le LLM Ouvrier (modif XML)
├── system_prompt_chef.md  ← Instructions pour le LLM Chef (SiaGPT, choix des tools)
├── skill/                 ← Documentation de référence (PAS dans Docker)
├── Dockerfile
├── requirements.txt
├── rebuild.sh             ← Script dev : rebuild Docker + relance
├── .env.example
└── .gitignore
```

### main.py (~960 lignes)

Le cœur du service. Contient :
- **Endpoints REST** : `/api/edit`, `/api/create`, `/api/generate`, `/api/inspect`
- **Serveur MCP** : tools `generate_pptx` et `edit_pptx` (transport SSE + Streamable HTTP)
- **Orchestration** : inspection → planification → modification XML → validation → repackage → upload
- **Fonctions core** : `_do_edit()` et `_do_create()` partagées entre REST et MCP

### pptx_tools.py (~540 lignes)

Manipulation PPTX pure. Zéro logique métier, zéro validation. Détaillé ci-dessus.

### pptx_validate.py (~680 lignes)

Validation complète en deux niveaux. Détaillé ci-dessus.

### schemas/ (~530 Ko)

Schemas XSD officiels de la norme Office Open XML (ISO/IEC 29500), copiés dans Docker pour la validation en runtime. Contient `pml.xsd` (PresentationML), `dml-main.xsd` (DrawingML), `opc-*.xsd` (packaging).

### system_prompt.md (~240 lignes)

Le "cahier des charges" du LLM Ouvrier. Définit les 2 phases (planification JSON + modification XML), le format XML PowerPoint, les bonnes pratiques et les guidelines de design. **C'est le levier principal pour améliorer la qualité des modifications XML.**

### system_prompt_chef.md (~100 lignes)

Les instructions pour le LLM Chef (celui de SiaGPT). Définit quand utiliser `generate_pptx` vs `edit_pptx`, comment choisir le bon template, comment rédiger un bon prompt, et quand poser des questions à l'utilisateur. **À copier dans la config du Chef (Langflow, system prompt SiaGPT, etc.)** Contient une section templates à remplir quand les templates Sia seront uploadés.

### skill/ — Documentation de référence

Contient le **skill PPTX original d'Anthropic** (celui que Claude utilise dans Cowork). **PAS copié dans Docker**, **PAS utilisé en runtime**. Les schemas et la logique de validation ont été extraits dans `schemas/` et `pptx_validate.py`. Reste dans le repo comme documentation pour les devs.

---

## Points d'entrée

### REST

| Endpoint | Méthode | Description |
|----------|---------|-------------|
| `/api/generate` | POST | Endpoint unifié — crée ou modifie selon présence d'un fichier |
| `/api/create` | POST | Créer un PPTX (depuis template ou squelette vierge) |
| `/api/edit` | POST | Modifier un PPTX existant (upload du fichier) |
| `/api/inspect` | POST | Structure JSON d'un PPTX |
| `/api/inspect/xml` | POST | XML brut d'une slide |
| `/health` | GET | Health check |

```bash
# Création sans template (squelette vierge)
curl -X POST http://localhost:8000/api/generate \
  -H "Content-Type: application/json" \
  -d '{"prompt": "Crée 5 slides sur l'\''IA en entreprise"}'

# Création avec template Sia Partners
curl -X POST http://localhost:8000/api/generate \
  -H "Content-Type: application/json" \
  -d '{"prompt": "Propale pour Airbus", "template_file_id": "abc-111-..."}'

# Édition d'un fichier existant (upload direct)
curl -X POST http://localhost:8000/api/edit \
  -F "prompt=Change tous les titres en bleu" \
  -F "file=@presentation.pptx"
```

### MCP (Model Context Protocol)

| Tool | Paramètres | Description |
|------|-----------|-------------|
| `generate_pptx` | `prompt`, `template_file_id`* | Crée un PPTX (depuis template ou squelette vierge), l'uploade |
| `edit_pptx` | `prompt`, `source_file_id` | Télécharge un PPTX existant, le modifie, l'uploade |

\* `template_file_id` est optionnel. Si fourni, le service télécharge le template depuis SiaGPT Medias et l'utilise comme base. Si omis, crée un squelette vierge (5 slides blanches).

**URL MCP** : `http://ADRESSE:8000/mcp/sse` (Streamable HTTP/SSE)

---

## Démarrage rapide

### 1. Configuration

```bash
cp .env.example .env
# Remplir LLM_API_KEY et SIAGPT_COLLECTION_ID
```

### 2. Docker

```bash
docker build -t pptx-service .
docker run -d -p 8000:8000 --env-file .env pptx-service
```

### 3. Vérification

```bash
curl http://localhost:8000/health
```

---

## Variables d'environnement

| Variable | Requis | Défaut | Description |
|----------|--------|--------|-------------|
| `LLM_API_KEY` | Oui | — | Bearer token SiaGPT |
| `SIAGPT_COLLECTION_ID` | Oui | — | UUID de la collection cible |
| `LLM_API_URL` | Non | `https://backend.siagpt.ai/chat/plain_llm` | URL de l'API LLM |
| `LLM_MODEL` | Non | `claude-4.5-sonnet` | Modèle LLM |
| `SIAGPT_MEDIAS_URL` | Non | `https://backend.siagpt.ai/medias` | URL API Medias |
| `MAX_RETRIES` | Non | `4` | Tentatives si XML invalide |

---

## Sécurité

Le service n'exécute **aucun code généré par le LLM**. Le LLM retourne uniquement du texte (JSON pour la planification, XML pour les modifications). Le service valide le XML avant de l'écrire.

---

## Limitations connues

- **Pas de QA visuelle** : pas de vérification du rendu (nécessiterait LibreOffice)
- **Pas de gestion d'images** : le LLM ne peut pas ajouter/modifier des images
- **Pas de graphiques/charts** : les graphiques Excel embarqués ne sont pas modifiables
- **Dépendance au modèle** : Claude Sonnet 4.5 donne de bons résultats, les modèles moins capables font plus d'erreurs XML

---

## Pour aller plus loin

- **Améliorer le system prompt** (`system_prompt.md`) : ajouter des exemples XML spécifiques aux templates Sia
- **QA visuelle** : si `/plain_llm` supporte les images, intégrer LibreOffice + validation visuelle
- **Templates pré-chargés** : bibliothèque de templates Sia Partners
- **Consulter `skill/`** : les scripts originaux contiennent des patterns avancés (images, thumbnails, PDF)
