import os
from typing import List, Dict, Any
from bs4 import BeautifulSoup
import docx

def parse_word_doc(file_path: str, category: str = "specification") -> List[Dict[str, Any]]:
    chunks = []
    fname = os.path.basename(file_path)
    try:
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        doc = docx.Document(file_path)
        current_section = fname
        current_paras = []

        for child in doc.element.body:
            if child.tag.endswith('p'):
                p = Paragraph(child, doc)
                text = p.text.strip()
                if not text:
                    continue
                if p.style and "heading" in p.style.name.lower():
                    if current_paras:
                        chunks.append({
                            "doc_title": f"{fname} - {current_section}",
                            "category": category,
                            "content": "\n".join(current_paras)
                        })
                        current_paras = []
                    current_section = text
                else:
                    current_paras.append(text)
            elif child.tag.endswith('tbl'):
                tbl = Table(child, doc)
                tbl_lines = []
                for idx, row in enumerate(tbl.rows):
                    row_cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    tbl_lines.append("| " + " | ".join(row_cells) + " |")
                    if idx == 0:
                        tbl_lines.append("| " + " | ".join(["---"] * len(row_cells)) + " |")
                if tbl_lines:
                    current_paras.append("\n" + "\n".join(tbl_lines) + "\n")

        if current_paras:
            chunks.append({
                "doc_title": f"{fname} - {current_section}",
                "category": category,
                "content": "\n".join(current_paras)
            })
    except Exception as e:
        print(f"Error reading Word doc {fname}: {e}")
    return chunks

def parse_html_mockup(file_path: str, category: str = "ui_design") -> List[Dict[str, Any]]:
    fname = os.path.basename(file_path)
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            soup = BeautifulSoup(f.read(), "html.parser")
            for tag in soup(["script", "style", "meta", "noscript"]):
                tag.decompose()
            text = re.sub(r'\s+', ' ', soup.get_text()).strip()
            return [{
                "doc_title": f"UI Spec: {fname}",
                "category": category,
                "content": text[:8000]
            }]
    except Exception as e:
        print(f"Error parsing HTML {fname}: {e}")
        return []
