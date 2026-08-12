import json
import sys
from pathlib import Path


def get_project_root():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    
    return Path(__file__).resolve().parent.parent


def load_config():
    project_root = get_project_root()
    config_file = project_root / "config.json"

    with open(config_file, "r", encoding="utf-8") as file:
        return json.load(file)