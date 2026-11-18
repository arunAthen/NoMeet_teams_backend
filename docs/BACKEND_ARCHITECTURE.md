# NoMeet Backend — Technical Architecture & API Specification

> **Component:** FastAPI Core Delivery Intelligence Service  
> **Repository:** `NoMeet_teams_backend`  
> **Runtime Environment:** Python 3.10+ / 3.12 on Vercel Serverless (Amazon Linux)  
> **Base URL:** `https://no-meet-teams-backend.vercel.app`  
> **Database:** SQLite 3 with Virtual FTS5 Full-Text Engine (`/tmp/teams_365_agent.db`)  

---

## 1. Architectural Philosophy & Overview

The NoMeet backend is designed as a **deterministic, evidence-first intelligence service**. Rather than operating as an unconstrained generative chatbot, the service functions as an autonomous reasoning layer over project tracking artifacts and live communications.

```
                              HTTP Request
                                   │
                                   ▼
             FastAPI Application Gateway (`app.main:app`)
                                   │
       ┌───────────────────────────┴───────────────────────────┐
       ▼                                                       ▼
TeamsSecurityHeadersMiddleware                        CORSMiddleware
(CSP frame-ancestors enforcement)             (Allowed origins, credentials, verbs)
       │                                                       │
       └───────────────────────────┬───────────────────────────┘
                                   │
                                   ▼
                     FastAPI REST Router (`/api/v1/teams`)
                                   │
       ┌───────────────────────────┼───────────────────────────┐
       ▼                           ▼                           ▼
Workspace & Sessions       Diagnostic & Status         ReAct Reasoning Engine
Management Endpoints            Endpoints               (`run_agent_query`)
       │                           │                           │
       ▼                           ▼                           ▼
SQLite Relational Tables    Azure Entra ID Client       Agent Tools & Search
(Items, Messages, Sessions) (`graph_client.py`)        (SQL Validator & FTS5 BM25)
       │                           │                           │
       └───────────────────────────┴───────────────────────────┘
                                   │
                                   ▼
                   Multi-Tier Model Inference Hierarchy
                   ├── Tier 1: Groq Cloud (Llama 3.3 70B)
                   ├── Tier 2: Google Gemini (1.5 Flash)
                   └── Tier 3: Deterministic Grounded Engine
```

---

## 2. Relational Database & FTS5 Full-Text Engine

The backend relies on an embedded SQLite database configured for high-concurrency read-heavy workloads. On Vercel serverless containers, the database dynamically provisions in `/tmp/teams_365_agent.db` with safe WAL/DELETE pragma switching.

### 2.1. Relational Tables Schema (`app/db/schema.py`)

#### 1. `project_items`
Stores structured work items (Change Requests, Bugs, Tasks, Milestones).
```sql
CREATE TABLE project_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL DEFAULT 'teams_delivery',
    ticket_id TEXT UNIQUE NOT NULL,
    item_type TEXT NOT NULL DEFAULT 'task', -- 'task', 'bug', 'change_request', 'milestone'
    sheet_name TEXT DEFAULT 'General',
    title TEXT NOT NULL,
    description TEXT DEFAULT '',
    module TEXT DEFAULT 'General',
    assignee TEXT DEFAULT 'Unassigned',
    priority TEXT DEFAULT 'Medium',
    status TEXT DEFAULT 'Open',
    planned_date TEXT,
    target_date TEXT,
    is_verified BOOLEAN DEFAULT 0,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);
CREATE INDEX idx_project_items_project ON project_items(project_id);
CREATE INDEX idx_project_items_ticket ON project_items(ticket_id);
CREATE INDEX idx_project_items_type ON project_items(item_type);
CREATE INDEX idx_project_items_status ON project_items(status);
```

#### 2. `teams_messages`
Stores indexed communications and signal records from Microsoft Teams chats and channels.
```sql
CREATE TABLE teams_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL DEFAULT 'teams_delivery',
    message_id TEXT UNIQUE NOT NULL,
    thread_id TEXT NOT NULL,
    sender_name TEXT NOT NULL,
    sender_email TEXT,
    timestamp TEXT NOT NULL,
    content TEXT NOT NULL,
    has_attachments BOOLEAN DEFAULT 0,
    is_technical_signal BOOLEAN DEFAULT 1,
    has_blocker BOOLEAN DEFAULT 0,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);
CREATE INDEX idx_teams_messages_thread ON teams_messages(thread_id);
CREATE INDEX idx_teams_messages_sender ON teams_messages(sender_name);
```

#### 3. `chat_sessions`
Represents user conversation sessions with strict AAD ID ring-fencing.
```sql
CREATE TABLE chat_sessions (
    session_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL DEFAULT 'default',
    user_aad_id TEXT DEFAULT 'default_user',
    user_name TEXT DEFAULT 'Team Member',
    title TEXT DEFAULT 'Conversation',
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);
CREATE INDEX idx_chat_sessions_user ON chat_sessions(user_aad_id);
CREATE INDEX idx_chat_sessions_project ON chat_sessions(project_id);
```

#### 4. `chat_messages`
Stores individual chronological turns within a session.
```sql
CREATE TABLE chat_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL, -- 'user', 'assistant'
    content TEXT NOT NULL,
    evidence TEXT,      -- Serialized JSON array of ReAct tool steps
    eval_scores TEXT,   -- Serialized JSON object of evaluation metrics
    created_at TIMESTAMP
);
CREATE INDEX idx_chat_messages_session ON chat_messages(session_id);
```

#### 5. `workspaces`
Tracks dynamically discovered and registered Teams group chats, channels, and projects.
```sql
CREATE TABLE workspaces (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    workspace_type TEXT DEFAULT 'chat', -- 'chat', 'channel', 'personal'
    last_active TIMESTAMP
);
```

### 2.2. SQLite Virtual FTS5 Full-Text Search Engine

For unstructured information retrieval, NoMeet initializes virtual tables compiled with the **Porter Stemmer** and **Unicode 6.1 tokenizer**:
```sql
CREATE VIRTUAL TABLE specs_fts USING fts5(
    doc_title, category, content, tokenize='porter unicode61'
);

CREATE VIRTUAL TABLE teams_fts USING fts5(
    message_id UNINDEXED, sender, timestamp UNINDEXED, content, tokenize='porter unicode61'
);

CREATE VIRTUAL TABLE items_fts USING fts5(
    ticket_id UNINDEXED, title, description, assignee, status, module, tokenize='porter unicode61'
);
```

#### Why FTS5 over External Vector Databases?
1. **Sub-Millisecond Query Times:** BM25 token matching executes locally inside the database engine in $< 1\text{ms}$.
2. **Zero Network Latency:** Eliminates external round-trips to Pinecone, Qdrant, or Weaviate.
3. **Exact Term Matching:** Guarantees exact identification of error codes, ticket IDs (`CR-012`), and technical keywords without embedding drift.

---

## 3. Dynamic Ingestion Engine

The ingestion pipeline transforms diverse project artifacts into normalized relational and full-text records.

### 3.1. Dynamic Excel Parser (`app/ingestion/dynamic_excel_parser.py`)
- **Arbitrary Column Layouts:** Scans sheets dynamically to identify header rows containing keywords like `Ticket`, `Status`, `Assignee`, `Module`, `Description`.
- **Flexible Type Casting:** Maps custom column labels to canonical schema fields.
- **Deduplication:** Uses SQLite `INSERT OR REPLACE` based on `ticket_id` to prevent duplicate ingestion across repeated sync cycles.

### 3.2. Word Specification Parser (`app/ingestion/docx_spec_parser.py`)
- **Document Chunking:** Traverses `.docx` XML trees, partitioning content by Heading 1 / Heading 2 boundaries.
- **Context Preservation:** Associates table structures and UI component descriptions with the parent document title before populating `specs_fts`.

---

## 4. Multi-Tier ReAct Reasoning Agent

The core delivery copilot is driven by an autonomous ReAct (Reason + Act) loop implemented in `app/agents/delivery_agent.py`.

### 4.1. The ReAct Tool Protocol
The agent is provided with two deterministic tools:
1. `execute_sql_query`:
   - Validates that incoming queries are strictly read-only `SELECT` or `WITH` operations.
   - Forbids modification keywords (`DROP`, `ALTER`, `DELETE`, `UPDATE`, `INSERT`).
   - Caps result sets at 20 rows to prevent token exhaustion.
2. `search_project_records`:
   - Executes BM25 ranking across `specs_fts` and `teams_fts`.
   - Returns highlighted snippets (`snippet(specs_fts, 2, '<b>', '</b>', '...', 25)`) with structured citation anchors.

### 4.2. Three-Tier Inference & High-Availability Hierarchy

```
                            User Query
                                │
                                ▼
         ┌──────────────────────────────────────────────┐
         │ Tier 1: Groq Cloud API (Llama 3.3 70B)       │
         │ - Sub-second latency via Groq LPU            │
         │ - Function calling / Tool execution          │
         └──────────────────────┬───────────────────────┘
                                │ (On Timeout / Rate Limit / Error)
                                ▼
         ┌──────────────────────────────────────────────┐
         │ Tier 2: Google Gemini (1.5 Flash)            │
         │ - High-context fallback                      │
         └──────────────────────┬───────────────────────┘
                                │ (On Quota / Network Downtime)
                                ▼
         ┌──────────────────────────────────────────────┐
         │ Tier 3: Deterministic Grounded Engine        │
         │ - 100% Offline Rule-Based SQL & FTS Parsing  │
         │ - Guarantees 200 OK with zero 500 crashes    │
         └──────────────────────────────────────────────┘
```

### 4.3. Real-Time Evaluation Engine (`app/agents/evaluator.py`)
Every assistant response is programmatically evaluated and assigned quality scores:
- **Faithfulness (0.00 – 1.00):** Measures whether every statement is backed by an explicit citation (`[Doc: ...]`, `[Tracker: ...]`, `[Teams: ...]`).
- **Answer Relevancy (0.00 – 1.00):** Measures query-response semantic alignment.
- **Groundedness (0.00 – 1.00):** Evaluates citation frequency and evidence density.

---

## 5. Microsoft Graph API & Azure Entra ID Integration

### 5.1. Authentication Architecture
- **Grant Type:** OAuth 2.0 Client Credentials Grant (`grant_type=client_credentials`).
- **Token URL:** `https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token`
- **Scope:** `https://graph.microsoft.com/.default`

### 5.2. Live Chat Sync Workflow (`app/integrations/teams/graph_client.py`)
1. Backend exchanges `AZURE_CLIENT_ID` and `AZURE_CLIENT_SECRET` for an application JWT token.
2. Queries `GET https://graph.microsoft.com/v1.0/chats/{chat_id}` to resolve the chat's friendly name (e.g., "Intranet Solutions").
3. Queries `GET https://graph.microsoft.com/v1.0/chats/{chat_id}/messages?$top=50` to pull live messages.
4. Strips HTML tags, sanitizes content, filters out attachment markers, and indexes messages into `teams_messages` and `teams_fts`.

---

## 6. Complete REST API Endpoint Reference

| HTTP Verb | Path | Request Body | Description | Response Status |
|---|---|---|---|---|
| `GET` | `/health` | None | Service & database health check | `200 OK` |
| `GET` | `/api/v1/teams/azure/status` | None | Validates Azure Entra ID token & Graph API connectivity | `200 OK` |
| `GET` | `/api/v1/teams/groups` | None | Lists all discovered workspaces (cleaned & decoded) | `200 OK` |
| `POST` | `/api/v1/teams/groups/register` | `RegisterWorkspaceRequest` | Auto-registers an active group chat or channel | `200 OK` |
| `GET` | `/api/v1/teams/sessions` | Query params: `user_id`, `group_id` | Lists past sessions strictly filtered by user AAD ID | `200 OK` |
| `POST` | `/api/v1/teams/sessions` | `CreateSessionRequest` | Creates a new user-isolated conversation session | `200 OK` |
| `GET` | `/api/v1/teams/sessions/{id}` | Path param: `session_id` | Retrieves message history for a specific session | `200 OK` |
| `DELETE` | `/api/v1/teams/sessions/{id}` | Path param: `session_id` | Deletes a conversation session and messages | `200 OK` |
| `POST` | `/api/v1/teams/query` | `QueryRequest` | Executes ReAct reasoning loop for a user query | `200 OK` |
| `GET` | `/api/v1/teams/overview` | Query param: `group_id` | Returns KPI metrics (CRs, bugs, indexed msgs) | `200 OK` |
| `POST` | `/api/v1/teams/sync/trigger` | Query param: `group_id` | Triggers file crawl and live chat synchronization | `200 OK` |
| `POST` | `/api/v1/teams/sync/chat` | `{"chat_id": "..."}` | Synchronizes messages for a specific chat ID | `200 OK` |
| `POST` | `/api/v1/teams/messages` | Bot Framework Activity | Handles 1:1 Bot Framework messages | `200 OK` |
| `POST` | `/api/v1/teams/webhook` | Graph Change Payload | Handles Microsoft Graph subscription webhook | `200 OK` |
