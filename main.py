"""BidBrief entry point.

Usage:
  python main.py                launch the GUI
  python main.py --cli DIR      process a folder from the command line
  python main.py --selftest     end-to-end selftest (generates a sample PDF,
                                calls the real API)
"""
import argparse
import os
import sys


def main():
    ap = argparse.ArgumentParser(description="BidBrief — 招标文件要点提取")
    ap.add_argument("--cli", metavar="目录", help="命令行模式：处理指定文件夹")
    ap.add_argument("--ext", default=".pdf,.docx,.xlsx,.xls,.ppt,.pptx,.txt,.md",
                    help="命令行模式：文件后缀，逗号分隔")
    ap.add_argument("--formats", default="xlsx,docx,pptx,md",
                    help="命令行模式：导出格式，逗号分隔（xlsx/docx/pptx/md）")
    ap.add_argument("--out", default=None, help="命令行模式：输出目录")
    ap.add_argument("--selftest", action="store_true", help="生成样例PDF并做端到端自测")
    args = ap.parse_args()

    if args.selftest:
        from selftest import run_selftest
        ok = run_selftest()
        sys.exit(0 if ok else 1)

    if args.cli:
        from bidbrief.config import load_config
        from bidbrief.pipeline import run_and_export
        from bidbrief.scanner import scan_folder

        cfg = load_config()
        if not cfg.get("api_key"):
            print("请先在 config.json 中填写 api_key")
            sys.exit(1)
        exts = [e.strip() for e in args.ext.split(",") if e.strip()]
        formats = [e.strip() for e in args.formats.split(",") if e.strip()]
        files = scan_folder(args.cli, exts, recursive=True)
        print(f"找到 {len(files)} 个文件")
        out = args.out or os.path.join(args.cli, "BidBrief_输出")
        result = run_and_export(cfg, files, out, formats=formats, on_log=print)
        n_ok = sum(1 for f in result["file_results"] if f["status"].startswith("完成"))
        print(f"完成：{n_ok}/{len(files)}")
        for k, v in result.get("exports", {}).items():
            print(f"  {k}: {v}")
        sys.exit(0)

    from bidbrief.ui import run_app
    run_app()


if __name__ == "__main__":
    main()
