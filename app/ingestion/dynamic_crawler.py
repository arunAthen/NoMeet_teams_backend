import os
import hashlib
import sqlite3
from pathlib import Path
from typing import Dict, Any, List
from app.core.config import settings
from app.ingestion.dynamic_excel_parser import parse_dynamic_excel
from app.ingestion.teams_chat_chunker import process_teams_chat_json
from app.ingestion.docx_spec_parser import parse_word_doc, parse_html_mockup

DB_FILE = settings.DATABASE_PATH

def compute_file_hash(path: Path) -> str:
    hasher = hashlib.md5()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()

def crawl_and_ingest_directory(root_dir: Path, project_id: str = "teams_delivery") -> Dict[str, int]:
    """Recursively walks any folder structure at any depth, handling any folder naming dynamically."""
    stats = {"excel_items": 0, "specs_indexed": 0, "chat_messages": 0, "files_scanned": 0}
    if not root_dir.exists():
        return stats

    with sqlite3.connect(DB_FILE) as conn:
        for file_path in root_dir.rglob("*"):
            if not file_path.is_file():
                continue
            stats["files_scanned"] += 1
            ext = file_path.suffix.lower()
            fname = file_path.name
            rel_path = str(file_path.relative_to(root_dir))
            rel_parts = file_path.relative_to(root_dir).parts
            file_project_id = rel_parts[0] if len(rel_parts) > 1 else project_id

            try:
                # 1. Excel Trackers (.xlsx)
                if ext == ".xlsx":
                    res = parse_dynamic_excel(str(file_path), project_id=file_project_id)
                    items = res.get("items", [])
                    specs = res.get("specs", [])
                    
                    for it in items:
                        conn.execute("""
                            INSERT INTO project_items (
                                project_id, ticket_id, item_type, sheet_name, title,
                                description, module, assignee, priority, status, planned_date, target_date, is_verified, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                            ON CONFLICT(ticket_id) DO UPDATE SET
                                title = excluded.title,
                                status = excluded.status,
                                assignee = excluded.assignee,
                                priority = excluded.priority,
                                updated_at = CURRENT_TIMESTAMP;
                        """, (
                            file_project_id, it["ticket_id"], it["item_type"], it["sheet_name"], it["title"],
                            it["description"], it["module"], it["assignee"], it["priority"], it["status"],
                            it["planned_date"], it["target_date"], 1 if it["is_verified"] else 0
                        ))
                        # Sync FTS
                        conn.execute("DELETE FROM items_fts WHERE ticket_id = ?", (it["ticket_id"],))
                        conn.execute("""
                            INSERT INTO items_fts (ticket_id, title, description, assignee, status, module)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (it["ticket_id"], it["title"], it["description"], it["assignee"], it["status"], it["module"]))
                        stats["excel_items"] += 1

                    for sp in specs:
                        conn.execute("DELETE FROM specs_fts WHERE doc_title = ?", (sp["doc_title"],))
                        conn.execute("""
                            INSERT INTO specs_fts (doc_title, category, content)
                            VALUES (?, ?, ?)
                        """, (sp["doc_title"], sp["category"], sp["content"]))
                        stats["specs_indexed"] += 1

                # 2. Teams Chat JSON (.json)
                elif ext == ".json" and ("teams" in fname.lower() or "chat" in fname.lower()):
                    c = process_teams_chat_json(str(file_path), project_id=file_project_id)
                    stats["chat_messages"] += c

                # 3. Word Documents (.docx)
                elif ext == ".docx":
                    chunks = parse_word_doc(str(file_path), category="specification")
                    for ck in chunks:
                        conn.execute("DELETE FROM specs_fts WHERE doc_title = ?", (ck["doc_title"],))
                        conn.execute("""
                            INSERT INTO specs_fts (doc_title, category, content)
                            VALUES (?, ?, ?)
                        """, (ck["doc_title"], ck["category"], ck["content"]))
                        stats["specs_indexed"] += 1

                # 4. HTML Mockups (.html)
                elif ext in (".html", ".htm"):
                    chunks = parse_html_mockup(str(file_path), category="ui_mockup")
                    for ck in chunks:
                        conn.execute("DELETE FROM specs_fts WHERE doc_title = ?", (ck["doc_title"],))
                        conn.execute("""
                            INSERT INTO specs_fts (doc_title, category, content)
                            VALUES (?, ?, ?)
                        """, (ck["doc_title"], ck["category"], ck["content"]))
                        stats["specs_indexed"] += 1

                conn.commit()
            except Exception as err:
                print(f"[Crawler error on {fname}]: {err}")

    return stats
