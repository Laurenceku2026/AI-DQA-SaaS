"""Fixed DFMEA template profiles for AI-DQA (deploy-safe standalone module)."""
from __future__ import annotations

import os
from typing import Any, Dict, Optional

MODULE_VERSION = "20260711"

TEMPLATE_PROFILES: Dict[str, Dict[str, Any]] = {
    "template1": {
        "filename": "模板1-DFMEA旧版.xlsx",
        "filename_en": "Template-1-DFMEA-Legacy.xlsx",
        "use_deepseek_fill": False,
        "use_deepseek_analysis": False,
        "label_zh": "模板1：DFMEA 旧版",
        "label_en": "Template 1: DFMEA Legacy",
    },
    "template2": {
        "filename": "模板2-DFMEA新版.xlsx",
        "filename_en": "Template-2-DFMEA-New.xlsx",
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
    "Template-1-DFMEA-Legacy.xlsx": "Template-1-DFMEA-Legacy.xlsx",
    "Template-2-DFMEA-New.xlsx": "Template-2-DFMEA-New.xlsx",
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


def profile_template_filename(mode: str, lang: str = "zh") -> Optional[str]:
    profile = TEMPLATE_PROFILES.get(mode)
    if not profile:
        return None
    if lang == "en":
        return str(profile.get("filename_en") or profile.get("filename") or "")
    return str(profile.get("filename") or "")


def resolve_profile_template_filename(mode: str, lang: str = "zh") -> Optional[str]:
    name = normalize_template_filename(profile_template_filename(mode, lang) or "")
    if not name:
        return None
    try:
        _resolve_template_path(name, "AI-DQA")
        return name
    except FileNotFoundError:
        if lang == "en":
            fallback = normalize_template_filename(profile_template_filename(mode, "zh") or "")
            if fallback and fallback != name:
                try:
                    _resolve_template_path(fallback, "AI-DQA")
                    return fallback
                except FileNotFoundError:
                    pass
        return name


def profile_uses_deepseek_fill(mode: str) -> bool:
    return bool(TEMPLATE_PROFILES.get(mode, {}).get("use_deepseek_fill"))


def profile_uses_deepseek_analysis(mode: str) -> bool:
    return bool(TEMPLATE_PROFILES.get(mode, {}).get("use_deepseek_analysis"))


def get_template_profile_label(mode: str, lang: str = "zh") -> str:
    profile = TEMPLATE_PROFILES.get(mode, {})
    key = "label_zh" if lang == "zh" else "label_en"
    return str(profile.get(key, mode))
