# NoMeet Teams Backend ⚡
> Autonomous Delivery & Operations Intelligence API for Microsoft Teams (FastAPI + Microsoft Graph + ReAct Copilot)

[![Production Status](https://img.shields.io/badge/Status-Live%20Production-success)](https://no-meet-teams-backend.vercel.app/health)
[![Vercel Deployment](https://img.shields.io/badge/Deployed%20on-Vercel-black)](https://no-meet-teams-backend.vercel.app)
[![API Version](https://img.shields.io/badge/Version-1.0.0-blue)](https://no-meet-teams-backend.vercel.app/docs)

`NoMeet` replaces synchronous delivery status meetings with an autonomous agent that ingests SharePoint files, crawls group chat messages, tracks Change Requests (CRs) and bugs, and responds to delivery questions with verifiable citations.

---

## 🌐 Live Production Endpoints

| Resource | URL | Description |
|---|---|---|
| **Health Check** | `https://no-meet-teams-backend.vercel.app/health` | SQLite database & API health verification |
| **Azure Status Diagnostic** | `https://no-meet-teams-backend.vercel.app/api/v1/teams/azure/status` | Live Entra ID token & Graph API permission status |
| **Interactive API Docs** | `https://no-meet-teams-backend.vercel.app/docs` | Swagger UI documentation |
| **Teams Personal Tab** | `https://no-meet-teams-backend.vercel.app/teams/tab/index.html` | Teams Cockpit Web interface |

---

## 📁 Repository Structure

```
NoMeet_teams_backend/
├── app/
│   ├── main.py                     # FastAPI application entrypoint & security middleware
│   ├── agents/
│   │   ├── delivery_agent.py       # Multi-tier ReAct delivery copilot reasoning engine
│   │   ├── agent_tools.py          # SQL & SQLite FTS5 BM25 retrieval tools
│   │   └── evaluator.py            # Groundedness & faithfulness mathematical scoring
│   ├── api/v1/
│   │   └── teams_routes.py         # REST endpoints for CRs, bugs, groups, chat, sync
│   ├── core/
│   │   ├── config.py               # Settings & environment configuration loader
│   │   └── security_headers.py     # Microsoft Teams iframe security middleware (CSP)
│   ├── db/
│   │   ├── schema.py               # SQLModel database models & SQLite tables
│   │   └── session.py              # SQLite session (with /tmp serverless migration)
│   ├── ingestion/
│   │   ├── dynamic_crawler.py      # Recursive folder crawler & ingestion orchestrator
│   │   ├── dynamic_excel_parser.py # Intelligent Excel CR & bug tracker parser
│   │   └── docx_parser.py          # Specification Word docx parser
│   ├── integrations/teams/
│   │   ├── graph_client.py         # Live Microsoft Graph API client & token diagnostics
│   │   ├── graph_delta_sync.py     # SharePoint file delta synchronization
│   │   ├── bot_adapter.py          # Microsoft Bot Framework activity adapter
│   │   ├── webhook_listener.py     # Graph change notification webhook listener
│   │   └── adaptive_cards.py       # Microsoft Teams Adaptive Cards generator
│   └── static/teams/               # Bundled Teams Tab web application
├── api/
│   └── index.py                    # Vercel Serverless function entrypoint
├── pyproject.toml                  # Python package configuration & dependencies
├── requirements.txt                # Pinned pip requirements
├── vercel.json                     # Vercel deployment configuration
├── .env.example                    # Environment variables template
└── README.md                       # Backend technical documentation
```

---

## ⚙️ Environment Variables Specification

The backend reads configuration from environment variables (or `.env` in local development):

| Variable | Required | Description |
|---|---|---|
| `GROQ_API_KEY` | Recommended | Groq API Key for Tier 1 LPU inference (`llama-3.3-70b-versatile`) |
| `GOOGLE_API_KEY` | Optional | Google Gemini API Key for Tier 2 fallback inference (`gemini-1.5-flash`) |
| `AZURE_CLIENT_ID` | Required for Live Sync | Azure Entra ID Application (Client) ID |
| `AZURE_CLIENT_SECRET`| Required for Live Sync | Azure Entra ID Client Secret Value |
| `AZURE_TENANT_ID` | Required for Live Sync | Azure Entra ID Directory (Tenant) ID |
| `AUTHORIZED_GROUP_CHAT_IDS` | Optional | Comma-separated list of Microsoft Teams Group Chat IDs |
| `PUBLIC_BASE_URL` | Optional | Public base URL for webhook callbacks |

---

## 🚀 Local Development Setup

### 1. Prerequisites
- Python 3.10+ installed

### 2. Install Dependencies
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Linux / macOS
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure Local Environment
```bash
cp .env.example .env
# Edit .env with your credentials
```

### 4. Run Development Server
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```
Access Swagger UI at `http://localhost:8000/docs`.

---

## 🛡️ Enterprise Security & Microsoft Graph Permissions

For live Microsoft Teams chat message and SharePoint file ingestion, configure the following **Application Permissions** in **Azure Portal (Microsoft Entra ID -> App Registrations)**:

1. `Chat.Read.All` (Application): To read group chat messages and technical discussions.
2. `ChannelMessage.Read.All` (Application): To index team channel conversations.
3. `Files.Read.All` (Application): To read project tracking Excel sheets and spec files.
4. `Sites.Read.All` (Application): To access SharePoint document libraries.

> ⚠️ **Admin Consent:** Application permissions require clicking **"Grant admin consent for [Organization]"** in the Azure Portal before Microsoft Graph endpoints become accessible. Test live status anytime via `/api/v1/teams/azure/status`.
