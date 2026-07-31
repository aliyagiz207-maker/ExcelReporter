import json
from pathlib import Path


def load_config():
    project_root = Path(__file__).resolve().parent.parent

    config_file = project_root / "config.json"

    with open(config_file, "r", encoding="utf-8") as file:
        return json.load(file)