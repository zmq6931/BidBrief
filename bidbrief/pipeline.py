"""Processing pipeline: scan -> parse -> chunk -> extract -> merge & locate pages -> export.

Supports cooperative cancellation throughout: the stop signal is checked
between every file and every chunk; on stop, results of completed files are
kept and exported as usual.
"""
import os
import threading

from . import exporter
from .chunker import chunk_document, locate_page
from .extractor import ApiAuthError, CancelledError, CATEGORIES, extract_chunk
from .parser import parse_file


class Pipeline:
    def __init__(self, cfg, cancel=None):
        self.cfg = cfg
        self.cancel = cancel or threading.Event()

    def stop(self):
        self.cancel.set()

    # ------------------------------------------------------------------
    def run(self, files, on_log=None, on_file=None, on_chunk=None):
        """Process a list of files.

        on_file(row, status, page_count, item_count)  status callback
        on_chunk(row, ci, total)                      chunk progress callback
        Returns {"file_results": [...], "cancelled": bool}.
        """
        results = []
        cancelled = False

        for row, path in enumerate(files):
            if self.cancel.is_set():
                cancelled = True
                break
            name = os.path.basename(path)
            self._log(on_log, f"开始处理：{name}")
            fr = {"path": path, "status": "", "page_count": None, "categories": {}}

            try:
                parsed = self._parse_step(fr, path, on_file, row, on_log)
                if parsed is None:
                    results.append(fr)
                    continue

                ok = self._extract_steps(fr, parsed, on_file, on_chunk, row, on_log)
                results.append(fr)
                if not ok:  # stopped mid-file
                    cancelled = True
                    break
            except ApiAuthError:
                fr["status"] = "失败(API认证)"
                self._notify_file(on_file, row, fr)
                results.append(fr)
                self._log(on_log, "API 认证失败，停止全部处理。请检查 config.json 中的 api_key。")
                cancelled = True
                break
            except CancelledError:
                fr["status"] = fr["status"] or "已停止"
                self._notify_file(on_file, row, fr)
                results.append(fr)
                cancelled = True
                break
            except Exception as e:
                fr["status"] = f"失败({e})"
                self._notify_file(on_file, row, fr)
                results.append(fr)
                self._log(on_log, f"处理失败：{name} -> {e}")

        return {"file_results": results, "cancelled": cancelled}

    # ------------------------------------------------------------------
    def _parse_step(self, fr, path, on_file, row, on_log=None):
        self._notify_status(on_file, row, fr, "解析中")
        parsed = parse_file(path)
        if not parsed["ok"]:
            fr["status"] = f"失败({parsed['error']})"
            self._notify_file(on_file, row, fr)
            self._log(on_log, f"解析失败：{os.path.basename(path)} -> {parsed['error']}")
            return None
        if parsed["is_scanned"]:
            fr["status"] = "跳过(扫描件需OCR)"
            self._notify_file(on_file, row, fr)
            return None
        fr["page_count"] = len(parsed["pages"]) if parsed.get("pages") is not None else None
        return parsed

    def _extract_steps(self, fr, parsed, on_file, on_chunk, row, on_log=None):
        name = os.path.basename(fr["path"])
        chunks = chunk_document(parsed, self.cfg.get("chunk_chars", 5000))
        if not chunks:
            fr["status"] = "完成(无文本内容)"
            self._notify_file(on_file, row, fr)
            return True

        merged = {c: [] for c in CATEGORIES}
        seen = set()  # dedupe by (category, quote prefix)

        for ci, chunk in enumerate(chunks, 1):
            if self.cancel.is_set():
                fr["status"] = "已停止(部分完成)" if any(merged.values()) else "已停止"
                fr["categories"] = self._locate(merged, parsed, chunks)
                self._notify_file(on_file, row, fr)
                return False
            if on_chunk:
                on_chunk(row, ci, len(chunks))
            sp, ep = chunk["start_page"], chunk["end_page"]
            loc = f"第{sp}-{ep}页" if sp is not None else f"块{ci}"
            self._notify_status(on_file, row, fr, f"抽取中({ci}/{len(chunks)}) {loc}")

            data = extract_chunk(self.cfg, chunk, self.cancel)
            for cat in CATEGORIES:
                for item in data.get(cat, []):
                    text = item.get("内容") or item.get("事项") or item.get("评分项") or ""
                    if not text and not item.get("原文摘录"):
                        continue
                    key = (cat, item.get("原文摘录", "")[:40])
                    if key in seen:
                        continue
                    seen.add(key)
                    item["_chunk_range"] = (sp, ep)
                    merged[cat].append(item)

        fr["categories"] = self._locate(merged, parsed, chunks)
        cnt = sum(len(v) for v in fr["categories"].values())
        fr["status"] = "完成(未发现要点)" if cnt == 0 else "完成"
        self._notify_file(on_file, row, fr)
        self._log(on_log, f"{name} 抽取完成，共 {cnt} 条要点")
        return True

    def _locate(self, merged, parsed, chunks):
        """Locate the exact page by searching the quote; fall back to the
        chunk's page range when the quote cannot be found."""
        pages = parsed.get("pages")
        for cat, items in merged.items():
            for item in items:
                sp, ep = item.pop("_chunk_range", (None, None))
                page = locate_page(pages, item.get("原文摘录", "")) if pages else None
                if page is not None:
                    item["页码"] = page
                elif sp is not None:
                    item["页码"] = f"{sp}-{ep}" if sp != ep else str(sp)
                else:
                    item["页码"] = None
        return merged

    # ------------------------------------------------------------------
    @staticmethod
    def _notify_status(on_file, row, fr, status):
        fr["status"] = status
        if on_file:
            on_file(row, fr["status"], fr.get("page_count"),
                    sum(len(v) for v in fr["categories"].values()))

    @staticmethod
    def _notify_file(on_file, row, fr):
        if on_file:
            on_file(row, fr["status"], fr.get("page_count"),
                    sum(len(v) for v in fr["categories"].values()))

    def _log(self, on_log, msg):
        if on_log:
            on_log(msg)


def run_and_export(cfg, files, out_dir, cancel=None, formats=None, on_log=None,
                   on_file=None, on_chunk=None):
    """Full flow: process then export in the requested formats. Returns the
    Pipeline.run result with export paths attached."""
    pipe = Pipeline(cfg, cancel)
    result = pipe.run(files, on_log=on_log, on_file=on_file, on_chunk=on_chunk)

    done = [fr for fr in result["file_results"]
            if fr.get("status", "").startswith("完成")]
    paths = {}
    if done:
        try:
            paths = exporter.export_all(done, out_dir, formats=formats)
            if on_log:
                for k, v in paths.items():
                    on_log(f"已导出：{v}")
        except Exception as e:
            if on_log:
                on_log(f"导出失败：{e}")
    elif on_log:
        on_log("没有成功处理的文件，跳过导出。")
    result["exports"] = paths
    return result
