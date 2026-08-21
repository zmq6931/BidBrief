"""Chunk document text and locate quotes on exact pages.

Chunking strategy: accumulate text page by page into chunks of roughly
chunk_chars; oversized single pages are split in place. Each chunk records its
page range; extracted items get an exact page by searching their quote back
into the per-page text.
"""
import re


def _clean(text):
    return text.strip()


def chunk_document(parsed, chunk_chars=5000):
    """Split a parse result into chunks.

    Returns [{"text", "start_page", "end_page"}] with 1-based page numbers;
    start_page/end_page are None when the document has no page info (DOCX/TXT).
    """
    pages = parsed.get("pages")
    text = parsed.get("text") or ""
    if pages is None:
        return _chunk_plain(text, chunk_chars)
    return _chunk_pages(pages, chunk_chars)


def _chunk_pages(pages, chunk_chars):
    chunks = []
    buf, buf_start, buf_end, size = [], None, 0, 0
    for i, raw in enumerate(pages, start=1):
        t = _clean(raw)
        if not t:
            continue
        # Oversized single page: split in-page, page number stays the same.
        if len(t) > chunk_chars * 1.5:
            if buf:
                chunks.append(_mk(buf, buf_start, buf_end))
                buf, buf_start, size = [], None, 0
            for j in range(0, len(t), chunk_chars):
                piece = t[j : j + chunk_chars]
                if piece.strip():
                    chunks.append(_mk([piece], i, i))
            buf_end = i
            continue
        if buf_start is None:
            buf_start = i
        buf.append(t)
        buf_end = i
        size += len(t)
        if size >= chunk_chars:
            chunks.append(_mk(buf, buf_start, buf_end))
            buf, buf_start, size = [], None, 0
    if buf:
        chunks.append(_mk(buf, buf_start, buf_end))
    return chunks


def _chunk_plain(text, chunk_chars):
    t = _clean(text)
    if not t:
        return []
    return [
        {"text": t[i : i + chunk_chars], "start_page": None, "end_page": None}
        for i in range(0, len(t), chunk_chars)
        if t[i : i + chunk_chars].strip()
    ]


def _mk(buf, start, end):
    return {"text": "\n".join(buf), "start_page": start, "end_page": end}


def _normalize(s):
    """Strip all whitespace (incl. full-width spaces and newlines) for loose
    matching against PDF-extracted text."""
    return re.sub(r"\s+", "", s or "")


def locate_page(pages, quote):
    """Find the 1-based page containing the quote; return None if not found."""
    q = _normalize(quote)
    if not q:
        return None
    normalized_pages = [_normalize(p) for p in (pages or [])]
    for probe_len in (30, 20, 12):
        probe = q[:probe_len]
        if len(probe) < min(probe_len, 8):
            continue
        for i, np in enumerate(normalized_pages, start=1):
            if probe in np:
                return i
    return None
