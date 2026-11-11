# NoMeet Teams Backend ⚡
> Autonomous Delivery & Operations Intelligence API for Microsoft Teams (FastAPI + Microsoft Graph + LLM Copilot)

`NoMeet` replaces synchronous delivery meetings with an autonomous agent that ingests SharePoint files, crawls group chat messages, tracks Change Requests (CRs) and bugs, and responds to delivery questions with grounded citations.

---

## 📁 Repository Structure
```
NoMeet_teams_backend/
├── app/
│   ├── main.py                     # FastAPI application entrypoint & middleware
│   ├── agents/
│   │   ├── delivery_agent.py       # LLM Delivery Copilot reasoning engine
│   │   ├── agent_tools.py          # Database & Graph retrieval tools
│   │   └── evaluator.py            # Answer evaluation & groundedness scoring
│   ├── api/v1/
│   │   └── teams_routes.py         # REST endpoints for CRs, bugs, groups, chat
│   ├── core/
│   │   ├── config.py               # Dynamic settings & environment loader
│   │   └── security_headers.py     # Microsoft Teams iframe security middleware
│   ├── db/
│   │   ├── schema.py               # SQLModel database models
│   │   └── session.py              # SQLite session (with /tmp serverless fallback)
│   ├── ingestion/
│   │   ├── dynamic_crawler.py      # Live Teams chat crawler
│   │   ├── dynamic_excel_parser.py # Excel CR/Bug parser
│   │   ├── docx_spec_parser.py     # Word spec & CR parser
│   │   └── teams_chat_chunker.py   # Chat message tokenizer & chunker
│   ├── integrations/teams/
│   │   ├── graph_delta_sync.py     # Microsoft Graph SharePoint file delta sync
│   │   ├── adaptive_cards.py       # Rich Teams Adaptive Cards
│   │   ├── bot_adapter.py          # Bot Framework adapter
│   │   └── webhook_listener.py     # Webhook listener
│   └── static/teams/               # Bundled Teams Tab fallback
├── api/
│   └── index.py                    # Vercel Serverless function entrypoint
├── data/
│   └── files_sync/                 # SharePoint sync directory (.gitkeep)
├── vercel.json                     # Vercel serverless configuration
├── requirements.txt                # Python dependencies
├── run_server.py                   # Local development server runner
├── Dockerfile                      # Production container configuration
├── pyproject.toml                  # Project metadata
├── .env.example                    # Environment variables template
├── .gitignore                      # Git ignore rules
└── README.md                       # Backend documentation
```

---

## 🚀 Quick Start (Local Development)

### 1. Prerequisites
- Python 3.10+ installed

### 2. Setup Environment
```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Configure `.env`
Copy `.env.example` to `.env` and fill in your keys:
```bash
cp .env.example .env
```
Key variables:
- `GROQ_API_KEY`: Groq API key for Llama 3.3 70B
- `AZURE_CLIENT_ID`: Azure Entra ID App Client ID
- `AZURE_CLIENT_SECRET`: Azure Entra ID App Client Secret
- `AZURE_TENANT_ID`: Azure Entra ID Tenant ID
- `AUTHORIZED_GROUP_CHAT_IDS`: Comma-separated list of Microsoft Teams Group Chat IDs

### 4. Start the Server
```bash
python run_server.py
```
- API Root: `http://localhost:8000/`
- Healthcheck: `http://localhost:8000/health`
- Interactive Swagger Docs: `http://localhost:8000/docs`
- Teams Personal Tab: `http://localhost:8000/teams/tab`

---

## ☁️ Deploying to Vercel

1. Push this repository to GitHub as **`NoMeet_teams_backend`**.
2. Go to [vercel.com](https://vercel.com) and click **Add New Project**.
3. Import the `NoMeet_teams_backend` GitHub repository.
4. In the **Environment Variables** section, add:
   - `GROQ_API_KEY`
   - `AZURE_CLIENT_ID`
   - `AZURE_CLIENT_SECRET`
   - `AZURE_TENANT_ID`
   - `AUTHORIZED_GROUP_CHAT_IDS`
5. Click **Deploy**. Vercel will build serverless handlers from `api/index.py` using `vercel.json`.

---

## 🛡️ Azure App Registration Setup

In **Microsoft Entra ID (Azure Portal)**:
1. Register a new application: `NoMeet Teams Agent`.
2. Under **API Permissions**, add the following **Application** permissions for Microsoft Graph:
   - `Files.Read.All` (To autonomously index SharePoint delivery specs and sheets)
   - `Chat.Read.WhereInstalled` (To read messages from chats where the app is added)
3. Grant **Admin Consent** for your tenant.
4. Under **Certificates & Secrets**, generate a new client secret and copy its value to `AZURE_CLIENT_SECRET`.

