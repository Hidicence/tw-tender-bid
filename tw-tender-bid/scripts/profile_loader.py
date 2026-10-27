# -*- coding: utf-8 -*-
"""
載入公司 profile。讀取順序:
  1. 環境變數 TW_TENDER_PROFILE 指向的 JSON
  2. ~/.claude/tender-profiles/default.json
  3. skill 內的 assets/company_profile.template.json(全是【待補】,審查會抓出來)

skill 本體不含任何公司資料;換公司使用時,只要換 profile。
"""
import json, os

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(SKILL_DIR, "assets", "company_profile.template.json")


def profile_path():
    env = os.environ.get("TW_TENDER_PROFILE")
    if env and os.path.exists(env):
        return env
    home = os.path.join(os.path.expanduser("~"), ".claude", "tender-profiles", "default.json")
    if os.path.exists(home):
        return home
    return TEMPLATE


def load():
    p = profile_path()
    with open(p, encoding="utf-8") as f:
        d = json.load(f)
    d["_path"] = p
    d["_is_template"] = (p == TEMPLATE)
    d.setdefault("brand_terms", [])
    d.setdefault("past_case_keywords", [])
    d.setdefault("seals_and_signature", {})
    d.setdefault("default_theme", "navy")
    return d


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    d = load()
    print("profile:", d["_path"], "| template" if d["_is_template"] else "| 已設定")
    print("公司:", d["name"])
