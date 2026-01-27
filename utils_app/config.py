import json
import os
from typing import Dict


def get_config_path() -> str:
    config_dir = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "GradeAnalysisSystem")
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "config.json")


def load_config() -> Dict:
    path = get_config_path()
    if not os.path.exists(path):
        data = {
            "api_key": "",
            "course_description": "",
            "objective_requirements": [],
            "previous_achievement_file": "",
            "report_style": "专业",
            "word_limit": 150,
            "ratios": {"usual": "", "midterm": "", "final": ""},
        }
        save_config(data)
        return data
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(data: Dict) -> None:
    path = get_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
