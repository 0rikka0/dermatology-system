# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Configuration Management
"""
import os
import json

CONFIG_FILE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")
DEFAULT_TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "模板")
DEFAULT_WORK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "工作目录")

DEFAULT_CONFIG = {
    "initialized": False,
    "work_dir": DEFAULT_WORK_DIR,
    "templates_dir": DEFAULT_TEMPLATES_DIR,
    "report_doctor": "张医生",
    "review_doctor": "李主任",
    "server_port": 8765,
    "auto_open_browser": True
}

def load_config():
    """Load configuration from config.json or return default."""
    if os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                config = DEFAULT_CONFIG.copy()
                config.update(data)
                return config
        except Exception as e:
            print(f"[Config] Error loading config.json: {e}")
    return DEFAULT_CONFIG.copy()

def save_config(config_dict):
    """Save configuration to config.json."""
    try:
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"[Config] Error saving config.json: {e}")
        return False

def update_config(updates):
    """Update specific configuration keys and save."""
    cfg = load_config()
    cfg.update(updates)
    save_config(cfg)
    return cfg
