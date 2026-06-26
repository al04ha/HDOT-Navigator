"""
chroma_store.py
All ChromaDB operations: indexing project files, indexing KB documents,
and querying both collections.
"""

import time
import chromadb

from config import CHROMA_STORAGE_DIR, COLLECTION_NAME, MANUAL_COLLECTION
from doc_extractor import extract_text_from_file, chunk_text
from ai_client import get_embedding

# Simple module-level singleton — no st.cache_resource needed
_client: chromadb.PersistentClient | None = None


def get_client() -> chromadb.PersistentClient:
    global _client
    if _client is None:
        _client = chromadb.PersistentClient(path=CHROMA_STORAGE_DIR)
    return _client


# ── Project file index ────────────────────────────────────────────────────────

def index_file_to_chroma(file_path: str, folder_display: str) -> None:
    """Index a project file into the main ChromaDB collection."""
    import os
    filename  = os.path.basename(file_path)
    raw_text  = extract_text_from_file(file_path)
    full_text = f"File Reference:\nFilename: {filename}\nFolder: {folder_display}\n\n{raw_text}"
    col = get_client().get_or_create_collection(name=COLLECTION_NAME)
    for idx, chunk in enumerate(chunk_text(full_text)):
        vec = get_embedding(chunk)
        if vec:
            col.add(
                embeddings=[vec],
                documents=[chunk],
                metadatas=[{"filename": filename, "folder": folder_display, "full_path": file_path}],
                ids=[f"id_{filename}_{idx}_{int(time.time())}"],
            )


def query_project_files(query: str, n_results: int = 6) -> tuple[list, list]:
    """Query the project file index. Returns (documents, metadatas)."""
    col = get_client().get_or_create_collection(name=COLLECTION_NAME)
    if col.count() == 0:
        return [], []
    vec = get_embedding(query)
    if not vec:
        return [], []
    res = col.query(query_embeddings=[vec], n_results=n_results)
    return res.get("documents", [[]])[0], res.get("metadatas", [[]])[0]


# ── Knowledge base index ──────────────────────────────────────────────────────

def index_text_to_kb(text: str, source_name: str, folder: str = "General") -> int:
    """Index raw text into the KB collection under source_name. Returns chunk count."""
    if not text.strip():
        return 0
    col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
    try:
        existing = col.get()
        old_ids = [existing["ids"][i] for i, m in enumerate(existing["metadatas"])
                   if m.get("source") == source_name]
        if old_ids:
            col.delete(ids=old_ids)
    except Exception:
        pass
    indexed = 0
    for idx, chunk in enumerate(chunk_text(text, chunk_size=500, overlap=100)):
        vec = get_embedding(chunk)
        if vec:
            col.add(
                embeddings=[vec],
                documents=[chunk],
                metadatas=[{"source": source_name, "folder": folder, "chunk": idx}],
                ids=[f"kb_{source_name}_{idx}_{int(time.time())}"],
            )
            indexed += 1
    return indexed


def index_file_to_kb(file_path: str, source_name: str, folder: str = "General") -> int:
    """Extract text from a file and index it into the KB."""
    return index_text_to_kb(extract_text_from_file(file_path), source_name, folder)


def is_source_indexed(source_name: str = "secretaries_manual") -> bool:
    """Check whether a given source name has been indexed in the KB."""
    try:
        col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
        existing = col.get(include=["metadatas"])
        return any(m.get("source") == source_name for m in existing.get("metadatas", []))
    except Exception:
        return False


def query_kb(query: str, n_results: int = 10, source_filter: str | None = None) -> tuple[list, list]:
    """Query the KB collection. Returns (documents, metadatas)."""
    try:
        col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
        if col.count() == 0:
            return [], []
        vec = get_embedding(query)
        if not vec:
            return [], []
        kwargs: dict = {"query_embeddings": [vec], "n_results": n_results}
        if source_filter:
            kwargs["where"] = {"source": source_filter}
        results = col.query(**kwargs)
        return results.get("documents", [[]])[0], results.get("metadatas", [[]])[0]
    except Exception:
        return [], []


def query_kb_by_folder(folder_name: str, n_results: int = 12) -> str:
    """Pull KB chunks from a specific folder and return them as joined text."""
    try:
        col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
        if col.count() == 0:
            return ""
        vec = get_embedding(f"document formatting template {folder_name}")
        if not vec:
            return ""
        results = col.query(
            query_embeddings=[vec],
            n_results=n_results,
            where={"folder": folder_name},
        )
        docs = results.get("documents", [[]])[0]
        return "\n---\n".join(docs)
    except Exception:
        return ""


def get_kb_all_metadata() -> list:
    """Return all metadata records from the KB collection."""
    try:
        col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
        if col.count() == 0:
            return []
        return col.get().get("metadatas", [])
    except Exception:
        return []


def kb_delete_by_source_and_folder(source: str, folder: str) -> int:
    """Delete all KB chunks matching a source+folder. Returns count deleted."""
    col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
    all_data = col.get()
    del_ids = [
        all_data["ids"][i]
        for i, m in enumerate(all_data["metadatas"])
        if m.get("source") == source and m.get("folder") == folder
    ]
    if del_ids:
        col.delete(ids=del_ids)
    return len(del_ids)


def kb_delete_by_folder(folder: str) -> int:
    """Delete all KB chunks in a folder. Returns count deleted."""
    col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
    all_data = col.get()
    del_ids = [
        all_data["ids"][i]
        for i, m in enumerate(all_data["metadatas"])
        if m.get("folder") == folder
    ]
    if del_ids:
        col.delete(ids=del_ids)
    return len(del_ids)


def kb_move_source_to_folder(source: str, from_folder: str, to_folder: str) -> int:
    """Move all chunks of a source from one folder to another. Returns count moved."""
    col = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
    all_data = col.get()
    moved = 0
    for i, m in enumerate(all_data["metadatas"]):
        if m.get("source") == source and m.get("folder") == from_folder:
            uid = all_data["ids"][i]
            col.update(ids=[uid], metadatas=[{**m, "folder": to_folder}])
            moved += 1
    return moved


def clear_all_collections() -> None:
    """Delete both ChromaDB collections entirely (used in Settings → Danger Zone)."""
    client = get_client()
    for name in [COLLECTION_NAME, MANUAL_COLLECTION]:
        try:
            client.delete_collection(name)
        except Exception:
            pass