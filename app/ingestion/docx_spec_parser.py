import os
from typing import List, Dict, Any
from bs4 import BeautifulSoup
import docx

def parse_word_doc(file_path: str, category: str = "requirements") -> List[Dict[str, Any]]:
    chunks = []
    fname = os.path.basename(file_path)
    try:
        doc = docx.Document(file_path)
        current_section = fname
        current_paras = []
        for p in doc.paragraphs:
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
