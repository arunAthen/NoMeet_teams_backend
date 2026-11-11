import logging
from typing import Dict
from pathlib import Path
from app.core.config import settings
from app.ingestion.dynamic_crawler import crawl_and_ingest_directory

logger = logging.getLogger(__name__)

def sync_sharepoint_files_delta() -> Dict[str, int]:
    """
    Syncs live files from the Teams Files folder (SharePoint).
    In production, connects via MS Graph API delta query (/drives/{id}/root/delta).
    In local/hybrid mode, monitors DATA_SYNC_DIR and SEED_DATA_DIR.
    """
    sync_dir = settings.DATA_SYNC_DIR
    seed_dir = settings.SEED_DATA_DIR

    total_stats = {"excel_items": 0, "specs_indexed": 0, "chat_messages": 0, "files_scanned": 0}

    # Ingest synced directory if populated
    if sync_dir.exists() and any(sync_dir.iterdir()):
        s1 = crawl_and_ingest_directory(sync_dir)
        for k in total_stats: total_stats[k] += s1.get(k, 0)

    # Ingest seed files directory
    if seed_dir.exists():
        s2 = crawl_and_ingest_directory(seed_dir)
        for k in total_stats: total_stats[k] += s2.get(k, 0)

    return total_stats
