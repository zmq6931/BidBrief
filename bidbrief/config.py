"""Load and save configuration (config.json).

When frozen with PyInstaller, config.json lives next to the exe so the API key
and custom prompt survive edits without repackaging.
"""
import json
import os
import sys

if getattr(sys, "frozen", False):  # packaged exe: use the exe's own folder
    _APP_DIR = os.path.dirname(sys.executable)
else:  # running from source: project root (parent of the bidbrief package)
    _APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CONFIG_FILE = os.path.join(_APP_DIR, "config.json")

DEFAULTS = {
    "api_key": "",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-chat",
    "custom_prompt": "",
    "max_tokens": 8192,
    "temperature": 0.1,
    "chunk_chars": 5000,
    "request_timeout": 180,
    "max_retries": 3,
    "retry_wait": 5,
}


def load_config():
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                user = json.load(f)
            if isinstance(user, dict):
                cfg.update(user)
        except Exception:
            pass
    return cfg


def save_config(cfg):
    merged = load_config()
    merged.update(cfg)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
