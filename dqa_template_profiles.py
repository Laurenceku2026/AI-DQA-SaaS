"""Fixed DFMEA template profiles for AI-DQA (deploy-safe standalone module)."""
from __future__ import annotations

import os
from typing import Any, Dict, Optional

TEMPLATE_PROFILES: Dict[str, Dict[str, Any]] = {
    "template1": {
        "filename": "模板1-DFMEA旧版.xlsx",
        "use_deepseek_fill": False,
        "use_deepseek_analysis": False,
        "label_zh": "模板1：DFMEA 旧版",
        "label_en": "Template 1: DFMEA Legacy",
    },
    "template2": {
        "filename": "模板2-DFMEA新版.xlsx",
        "use_deepseek_fill": True,
        "use_deepseek_analysis": True,
        "label_zh": "模板2：DFMEA 新版",
        "label_en": "Template 2: DFMEA New",
    },
}

# 兼容旧部署/旧配置中的模板文件名
TEMPLATE_FILENAME_ALIASES = {
    "DFMEA模板1.xlsx": "模板1-DFMEA旧版.xlsx",
    "新版FMEA表格.xlsx": "模板2-DFMEA新版.xlsx",
}


def normalize_template_filename(filename: str) -> str:
    return TEMPLATE_FILENAME_ALIASES.get(filename, filename)


def _resolve_template_path(filename: str, app_key: str = "AI-DQA") -> str:
    filename = normalize_template_filename(filename)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "templates", filename),
        os.path.join(os.environ.get("DFSS_TEMPLATE_DIR", ""), app_key, filename),
        os.path.join(r"C:\Users\Laurence\Technical\Project\SaaS\DFSS Report Template", app_key, filename),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    raise FileNotFoundError(f"Template not found: {filename}")


def resolve_profile_template_filename(mode: str) -> Optional[str]:
    profile = TEMPLATE_PROFILES.get(mode)
    if not profile:
        return None
    name = normalize_template_filename(profile.get("filename") or "")
    if not name:
        return None
    try:
        _resolve_template_path(name, "AI-DQA")
        return name
    except FileNotFoundError:
        return name


def profile_uses_deepseek_fill(mode: str) -> bool:
    return bool(TEMPLATE_PROFILES.get(mode, {}).get("use_deepseek_fill"))


def profile_uses_deepseek_analysis(mode: str) -> bool:
    return bool(TEMPLATE_PROFILES.get(mode, {}).get("use_deepseek_analysis"))


def get_template_profile_label(mode: str, lang: str = "zh") -> str:
    profile = TEMPLATE_PROFILES.get(mode, {})
    key = "label_zh" if lang == "zh" else "label_en"
    return str(profile.get(key, mode))
