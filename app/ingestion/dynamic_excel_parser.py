import os
import re
from datetime import datetime, date
from typing import List, Dict, Any, Optional
import openpyxl

ALIAS_TITLE = ["task", "description", "summary", "title", "requirement", "work item", "workitem", "defect", "topic", "query"]
ALIAS_ASSIGNEE = ["owner", "assignee", "assigned to", "assigned_to", "developer", "dev", "resource", "lead", "name", "raised by"]
ALIAS_STATUS = ["status", "state", "progress", "stage", "condition", "resolution"]
ALIAS_PRIORITY = ["priority", "severity", "urgency", "impact"]
ALIAS_MODULE = ["module", "component", "area", "phase", "epic"]
ALIAS_ID = ["bug id", "task id", "ticket", "s. no.", "s.no", "id", "key", "ref"]
ALIAS_DATE = ["target date", "due date", "planned start", "planned end", "raised on", "resolved on", "date", "eta"]

def normalize_date_val(val: Any) -> Optional[str]:
    if val is None or str(val).strip() == "":
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return val_str[:10] if len(val_str) >= 10 else val_str

def find_header_row(rows: List[tuple]) -> Tuple[int, List[str]]:
    """Inspects top 6 rows to dynamically detect the true header row."""
    best_row_idx = 0
    best_score = -1
    best_headers = []
    
    for idx, r in enumerate(rows[:6]):
        non_empty = [str(c).strip().lower() for c in r if c is not None and str(c).strip()]
        if not non_empty:
            continue
        # Count semantic matches
        score = sum(1 for c in non_empty if any(alias in c for alias in ALIAS_TITLE + ALIAS_ASSIGNEE + ALIAS_STATUS + ALIAS_ID))
        if score > best_score:
            best_score = score
            best_row_idx = idx
            best_headers = [str(c).strip() if c is not None else f"Col_{c_idx+1}" for c_idx, c in enumerate(r)]
            
    return best_row_idx, best_headers

def parse_dynamic_excel(file_path: str, project_id: str = "teams_delivery") -> Dict[str, Any]:
    """Universal Excel parser supporting arbitrary sheets and column aliases with zero crashes."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    file_name = os.path.basename(file_path)
    structured_items: List[Dict[str, Any]] = []
    unstructured_chunks: List[Dict[str, Any]] = []

    for sname in wb.sheetnames:
        ws = wb[sname]
        all_rows = list(ws.iter_rows(values_only=True))
        if len(all_rows) < 2:
            continue

        header_idx, headers = find_header_row(all_rows)
        if not headers:
            continue

        lower_headers = [h.lower() for h in headers]
        col_map: Dict[str, int] = {}
        for c_idx, h in enumerate(lower_headers):
            if any(a in h for a in ALIAS_ID) and "id" not in col_map:
                col_map["id"] = c_idx
            elif any(a in h for a in ALIAS_TITLE) and "title" not in col_map:
                col_map["title"] = c_idx
            elif any(a in h for a in ALIAS_ASSIGNEE) and "assignee" not in col_map:
                col_map["assignee"] = c_idx
            elif any(a in h for a in ALIAS_STATUS) and "status" not in col_map:
                col_map["status"] = c_idx
            elif any(a in h for a in ALIAS_PRIORITY) and "priority" not in col_map:
                col_map["priority"] = c_idx
            elif any(a in h for a in ALIAS_MODULE) and "module" not in col_map:
                col_map["module"] = c_idx
            elif any(a in h for a in ALIAS_DATE) and "date" not in col_map:
                col_map["date"] = c_idx

        # If a sheet lacks a recognizable title/task column, treat it as a specification matrix
        if "title" not in col_map:
            table_lines = [" | ".join(headers), " | ".join(["---"] * len(headers))]
            for r in all_rows[header_idx + 1 : 120]:
                if not any(r):
                    continue
                vals = [str(c).strip().replace("\n", " ") if c is not None else "" for c in r]
                table_lines.append(" | ".join(vals))
            if len(table_lines) > 2:
                unstructured_chunks.append({
                    "doc_title": f"Spreadsheet Spec: {file_name} - {sname}",
                    "category": "design_spec",
                    "content": "\n".join(table_lines)
                })
            continue

        # Extract structured items
        sheet_lower = sname.lower()
        default_type = "bug" if "bug" in sheet_lower or "defect" in sheet_lower else "task"

        for row_num, r in enumerate(all_rows[header_idx + 1:], start=header_idx + 2):
            if not any(r):
                continue
            title_val = r[col_map["title"]] if col_map["title"] < len(r) else None
            if not title_val or str(title_val).strip() == "":
                continue

            raw_id = r[col_map["id"]] if "id" in col_map and col_map["id"] < len(r) else row_num
            assignee = r[col_map["assignee"]] if "assignee" in col_map and col_map["assignee"] < len(r) else "Unassigned"
            status = r[col_map["status"]] if "status" in col_map and col_map["status"] < len(r) else "Open"
            priority = r[col_map["priority"]] if "priority" in col_map and col_map["priority"] < len(r) else "Medium"
            module = r[col_map["module"]] if "module" in col_map and col_map["module"] < len(r) else "General"
            date_val = normalize_date_val(r[col_map["date"]]) if "date" in col_map and col_map["date"] < len(r) else None

            # Detect Change Requests dynamically
            combined_text = f"{sname} {title_val} {module}".lower()
            is_cr = bool(re.search(r'\b(?:cr|change request)\b', combined_text))
            item_type = "change_request" if is_cr else default_type

            prefix = "BUG" if item_type == "bug" else ("CR" if item_type == "change_request" else "TSK")
            try:
                ticket_id = f"{prefix}-{int(raw_id):03d}" if str(raw_id).isdigit() else f"{prefix}-{str(raw_id).strip()}"
            except Exception:
                ticket_id = f"{prefix}-{row_num:03d}"

            structured_items.append({
                "project_id": project_id,
                "ticket_id": ticket_id,
                "item_type": item_type,
                "sheet_name": sname,
                "title": str(title_val).strip(),
                "description": f"Source: {file_name} [{sname}]. Module: {module}.",
                "module": str(module).strip() if module else "General",
                "assignee": str(assignee).strip() if assignee else "Unassigned",
                "priority": str(priority).strip() if priority else "Medium",
                "status": str(status).strip() if status else "Open",
                "planned_date": date_val,
                "target_date": date_val,
                "is_verified": str(status).strip().lower() in ["completed", "closed", "resolved", "verified", "done"]
            })

    return {
        "items": structured_items,
        "specs": unstructured_chunks
    }
