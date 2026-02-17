# PPTX Service

Micro-service de génération et modification de fichiers PowerPoint via LLM.
Conçu pour s'intégrer à SiaGPT via MCP (Model Context Protocol) ou API REST.

## Architecture

```
SiaGPT / Langflow → [MCP SSE] → PPTX Service → [API LLM] → SiaGPT (claude-4.5-sonnet)
                                       │
                                       ▼
                                 Exécute le code Python
                                 (édition XML ou python-pptx)
                                       │
                                       ▼
                                 Upload dans la collection SiaGPT
```

## Démarrage rapide

### 1. Configuration

```bash
cp .env.example .env
# Édite .env avec ton token et ton UUID de collection
```

### 2. Avec Docker

```bash
docker build -t pptx-service .
docker run -d -p 8000:8000 \
  -e LLM_API_URL=https://backend.siagpt.ai/chat/plain_llm \
  -e LLM_API_KEY=ton-bearer-token \
  -e LLM_MODEL=claude-4.5-sonnet \
  -e SIAGPT_COLLECTION_ID=uuid-de-ta-collection \
  pptx-service
```

### 3. Vérification

```bash
curl http://localhost:8000/health
```

## Intégration MCP (Langflow / SiaGPT)

Le service expose un endpoint MCP Streamable HTTP :

- **URL MCP** : `http://ADRESSE:8000/mcp/sse`
- **Transport** : Streamable HTTP/SSE
- **Tool exposé** : `generate_pptx` — paramètre `prompt` (string)

Test :
```bash
curl -X POST http://localhost:8000/mcp/sse \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

## API REST

| Endpoint | Méthode | Description |
|----------|---------|-------------|
| `/api/generate` | POST | Endpoint unifié (création ou édition) |
| `/api/create` | POST | Créer un PPTX from scratch |
| `/api/edit` | POST | Modifier un PPTX existant |
| `/api/inspect` | POST | Inspecter la structure d'un PPTX |
| `/health` | GET | Health check |

Exemples :
```bash
# Création
curl -X POST http://localhost:8000/api/create \
  -F "prompt=Crée 3 slides sur l'IA en entreprise"

# Édition
curl -X POST http://localhost:8000/api/edit \
  -F "prompt=Change les titres en bleu" \
  -F "file=@presentation.pptx"
```

## Variables d'environnement

| Variable | Requis | Défaut |
|----------|--------|--------|
| `LLM_API_KEY` | Oui | — |
| `SIAGPT_COLLECTION_ID` | Oui | — |
| `LLM_API_URL` | Non | `https://backend.siagpt.ai/chat/plain_llm` |
| `LLM_MODEL` | Non | `claude-4.5-sonnet` |
| `MAX_RETRIES` | Non | `4` |

## Sécurité

Le service utilise `exec()` pour exécuter le code généré par le LLM dans un container Docker isolé.
Pour la production : containers éphémères par requête, whitelist d'imports, timeout d'exécution.
