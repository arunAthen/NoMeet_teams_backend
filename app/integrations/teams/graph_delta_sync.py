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

    # Ingest seed files directory if present
    if seed_dir.exists():
        s2 = crawl_and_ingest_directory(seed_dir)
        for k in total_stats: total_stats[k] += s2.get(k, 0)

    # If synced directory is empty on cold start and Azure credentials exist, trigger live sync from Graph API
    if total_stats["specs_indexed"] == 0 and settings.AZURE_CLIENT_ID:
        try:
            import asyncio
            from app.integrations.teams.graph_client import sync_live_chat_messages
            target_chats = [
                "19:2337de30cb834d059c344292c8377656@thread.v2",
                "19:7bdcc4616332483b82e7488c4424ec23@thread.v2"
            ]
            if settings.AUTHORIZED_GROUP_CHAT_IDS:
                for cid in settings.AUTHORIZED_GROUP_CHAT_IDS:
                    if cid and cid not in target_chats:
                        target_chats.append(cid)

            async def _run_initial_sync():
                for cid in target_chats:
                    try:
                        await sync_live_chat_messages(cid)
                    except Exception as ce:
                        logger.warning(f"Initial sync warning for {cid}: {ce}")

            try:
                loop = asyncio.get_event_loop()
                if loop.is_running():
                    import concurrent.futures
                    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
                        pool.submit(asyncio.run, _run_initial_sync()).result(timeout=15)
                else:
                    loop.run_until_complete(_run_initial_sync())
            except Exception:
                asyncio.run(_run_initial_sync())

            if sync_dir.exists() and any(sync_dir.iterdir()):
                s3 = crawl_and_ingest_directory(sync_dir)
                for k in total_stats: total_stats[k] += s3.get(k, 0)
        except Exception as live_err:
            logger.warning(f"Live startup Graph sync notice: {live_err}")

    return total_stats
