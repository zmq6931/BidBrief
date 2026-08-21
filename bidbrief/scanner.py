"""Scan a folder for files matching selected extensions."""
import os

# Supported file types (keep in sync with parser.py)
SUPPORTED_EXTENSIONS = [".pdf", ".docx", ".xlsx", ".xls", ".ppt", ".pptx", ".txt", ".md"]


def scan_folder(folder, extensions, recursive=True):
    """Return matching file paths sorted by file name.

    extensions looks like [".pdf", ".docx"]; matching is case-insensitive.
    """
    exts = {e.lower() if e.startswith(".") else "." + e.lower() for e in extensions}
    files = []
    if recursive:
        for root, _dirs, names in os.walk(folder):
            for n in names:
                if os.path.splitext(n)[1].lower() in exts:
                    files.append(os.path.join(root, n))
    else:
        for n in os.listdir(folder):
            p = os.path.join(folder, n)
            if os.path.isfile(p) and os.path.splitext(n)[1].lower() in exts:
                files.append(p)
    return sorted(files, key=str.lower)
