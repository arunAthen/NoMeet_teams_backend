import os
from datetime import datetime, timezone
from typing import Optional, List
from sqlmodel import SQLModel, Field

def utc_now() -> datetime:
    return datetime.now(timezone.utc)

class ProjectItem(SQLModel, table=True):
    __tablename__ = "project_items"
    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: str = Field(default="teams_delivery", index=True)
    ticket_id: str = Field(index=True, unique=True)
    item_type: str = Field(default="task", index=True)  # task, bug, change_request, milestone
    sheet_name: str = Field(default="General")
    title: str
    description: str = Field(default="")
    module: str = Field(default="General", index=True)
    assignee: str = Field(default="Unassigned", index=True)
    priority: str = Field(default="Medium", index=True)
    status: str = Field(default="Open", index=True)
    planned_date: Optional[str] = None
    target_date: Optional[str] = None
    is_verified: bool = Field(default=False)
    created_at: Optional[datetime] = Field(default_factory=utc_now)
    updated_at: Optional[datetime] = Field(default_factory=utc_now)

class TeamsMessage(SQLModel, table=True):
    __tablename__ = "teams_messages"
    id: Optional[int] = Field(default=None, primary_key=True)
    project_id: str = Field(default="teams_delivery", index=True)
    message_id: str = Field(index=True, unique=True)
    thread_id: str = Field(index=True)
    sender_name: str = Field(index=True)
    sender_email: Optional[str] = None
    timestamp: str = Field(index=True)
    content: str
    has_attachments: bool = Field(default=False)
    is_technical_signal: bool = Field(default=True, index=True)
    has_blocker: bool = Field(default=False, index=True)
    created_at: Optional[datetime] = Field(default_factory=utc_now)
    updated_at: Optional[datetime] = Field(default_factory=utc_now)

class ChatSession(SQLModel, table=True):
    __tablename__ = "chat_sessions"
    session_id: str = Field(primary_key=True)
    project_id: str = Field(default="default", index=True)
    user_aad_id: Optional[str] = Field(default="default_user", index=True)
    user_name: Optional[str] = Field(default="Team Member")
    title: str = Field(default="Conversation")
    created_at: Optional[datetime] = Field(default_factory=utc_now)
    updated_at: Optional[datetime] = Field(default_factory=utc_now)

class ChatMessage(SQLModel, table=True):
    __tablename__ = "chat_messages"
    id: Optional[int] = Field(default=None, primary_key=True)
    session_id: str = Field(index=True)
    role: str = Field(index=True)  # user, assistant
    content: str
    evidence: Optional[str] = None  # JSON string
    eval_scores: Optional[str] = None  # JSON string
    created_at: Optional[datetime] = Field(default_factory=utc_now)

class DataSource(SQLModel, table=True):
    __tablename__ = "data_sources"
    id: Optional[int] = Field(default=None, primary_key=True)
    name: str = Field(unique=True)
    source_type: str  # excel, docx, teams_chat, html
    category: str  # requirements, design, testing, communications
    file_path: str
    file_hash: Optional[str] = None
    item_count: int = Field(default=0)
    last_synced: Optional[datetime] = Field(default_factory=utc_now)
    health_status: str = Field(default="healthy")

class AuditEvent(SQLModel, table=True):
    __tablename__ = "audit_events"
    id: Optional[int] = Field(default=None, primary_key=True)
    action: str
    entity_type: str
    entity_id: Optional[str] = None
    actor: str = Field(default="system")
    details: Optional[str] = None
    timestamp: Optional[datetime] = Field(default_factory=utc_now)
