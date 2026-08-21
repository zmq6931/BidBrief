"""Scan a folder for files matching selected extensions."""
import os

# Supported file types (keep in sync with parser.py)
SUPPORTED_EXTENSIONS = [".pdf", ".docx", ".xlsx", ".xls", ".ppt", ".pptx", ".txt", ".md"]


def scan_folder(folder, extensions, recursive=True, exclude_dirs=()):
    """Return matching file paths sorted by file name.

    extensions looks like [".pdf", ".docx"]; matching is case-insensitive.
    exclude_dirs (absolute paths) are skipped entirely - used to keep the
    tool from re-ingesting its own output folder on repeat runs.
    """
    exts = {e.lower() if e.startswith(".") else "." + e.lower() for e in extensions}
    excluded = {os.path.abspath(d).lower() for d in exclude_dirs if d}
    files = []
    if recursive:
        for root, dirs, names in os.walk(folder):
            abs_root = os.path.abspath(root).lower()
            # prune excluded directories from descent
            dirs[:] = [d for d in dirs
                       if os.path.abspath(os.path.join(root, d)).lower() not in excluded]
            if abs_root in excluded:
                continue
            for n in names:
                if os.path.splitext(n)[1].lower() in exts:
                    files.append(os.path.join(root, n))
    else:
        for n in os.listdir(folder):
            p = os.path.join(folder, n)
            if os.path.isfile(p) and os.path.splitext(n)[1].lower() in exts:
                files.append(p)
    return sorted(files, key=str.lower)
