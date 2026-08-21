"""Parse files (PDF / DOCX / TXT) into text; PDFs keep per-page text for page-level citation."""
import os


def parse_file(path):
    """Parse a single file.

    Returns a dict:
      ok         whether text was successfully extracted
      pages      list of per-page texts for PDFs; None for DOCX/TXT (no page info)
      text       full document text
      is_scanned PDF looks like a pure scan (no text layer)
      error      failure reason (set when ok is False)
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            return _parse_pdf(path)
        if ext == ".docx":
            return _parse_docx(path)
        if ext in (".txt", ".md"):
            return _parse_text(path)
        return _fail(f"不支持的文件格式: {ext}")
    except Exception as e:
        return _fail(f"解析失败: {e}")


def _fail(msg, **extra):
    d = {"ok": False, "pages": None, "text": "", "is_scanned": False, "error": msg}
    d.update(extra)
    return d


def _parse_pdf(path):
    import pymupdf  # PyMuPDF

    doc = pymupdf.open(path)
    if doc.needs_pass:
        if not doc.authenticate(""):
            doc.close()
            return _fail("PDF 已加密，无法读取（部分投标平台下载的文件带 CA 加密）")
    pages = []
    for page in doc:
        try:
            pages.append(page.get_text("text") or "")
        except Exception:
            pages.append("")
    doc.close()

    total = sum(len(p.strip()) for p in pages)
    n = max(len(pages), 1)
    if total / n < 50:
        # Almost no text layer: treat the whole file as a scan.
        return {
            "ok": True,
            "pages": pages,
            "text": "\n".join(pages),
            "is_scanned": True,
            "error": None,
        }
    return {"ok": True, "pages": pages, "text": "\n".join(pages), "is_scanned": False, "error": None}


def _parse_docx(path):
    from docx import Document

    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:  # scoring criteria often live inside tables
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells]
            line = " | ".join(x for x in cells if x)
            if line:
                parts.append(line)
    text = "\n".join(x for x in parts if x.strip())
    if not text.strip():
        return _fail("DOCX 中未提取到文本（可能为图片内容）")
    return {"ok": True, "pages": None, "text": text, "is_scanned": False, "error": None}


def _parse_text(path):
    text = ""
    for enc in ("utf-8", "gb18030"):
        try:
            with open(path, "r", encoding=enc) as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue
    if not text.strip():
        return _fail("文件为空或无法解码")
    return {"ok": True, "pages": None, "text": text, "is_scanned": False, "error": None}
