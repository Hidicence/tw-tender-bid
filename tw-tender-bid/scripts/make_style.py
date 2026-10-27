# -*- coding: utf-8 -*-
"""
產生新的風格包(assets/styles/*.json),讓每一案的風格可以不同。

來源一(預設):自己指定主色等色碼
    python -I make_style.py --id my-style --name 名稱 --primary 1F4E79 --accent E07B39 [--banner sidebar]
來源二:從本機 ui-ux-pro-max skill 的配色資料(colors.csv)挑一列轉換(選用,沒有該 skill 就用來源一)
    python -I make_style.py --from-uiux "Healthcare" --id med-calm --name 醫療平靜 [--uiux-data <資料夾>]
    python -I make_style.py --from-uiux "Healthcare" --list        只列出符合的配色列

規則:
  - 主色必須能承載白字(對比 >= 4.5),不夠就自動加深。
  - 色票是風格,字體、字級、紙張、邊界等「規格」不在這裡,由招標文件的 format.json 決定。
  - 若招標機關有自己的識別色或 CI 手冊,先用機關的顏色。
"""
import argparse, colorsys, csv, json, os, sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(SKILL_DIR, "assets", "styles")


def _hex(h):
    h = h.strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return h.upper()


def _lum(h):
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast_with_white(h):
    return 1.05 / (_lum(h) + 0.05)


def darken_until(h, ratio=4.5):
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    hh, l, s = colorsys.rgb_to_hls(r, g, b)
    while contrast_with_white(h) < ratio and l > 0.05:
        l -= 0.02
        r, g, b = colorsys.hls_to_rgb(hh, l, s)
        h = "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
    return h


def mix(h1, h2, t):
    a = [int(h1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(h2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join("%02X" % round(a[i] * (1 - t) + b[i] * t) for i in range(3))


def build_pack(pid, name, primary, accent, secondary=None, tint=None, school="自訂", mood="", best_for="", banner="filled",
               table="zebra", cover="left", fonts=("微軟正黑體", "微軟正黑體", "Calibri"), inspiration=""):
    primary = darken_until(_hex(primary))
    accent = _hex(accent)
    secondary = darken_until(_hex(secondary), 4.0) if secondary else mix(primary, "FFFFFF", 0.2)
    tertiary = mix(primary, accent, 0.45)
    tint = _hex(tint) if tint else mix(primary, "FFFFFF", 0.92)
    zebra = mix(primary, "FFFFFF", 0.96)
    chart = ["#" + x for x in (primary, secondary, tertiary, accent, mix(primary, "FFFFFF", 0.55), "718096", mix(primary, "000000", 0.4))]
    return {
        "id": pid, "name": name, "school": school, "mood": mood, "best_for": best_for, "avoid_for": "", "inspiration": inspiration,
        "colors": {"primary": primary, "secondary": secondary, "tertiary": tertiary, "accent": accent, "tint": tint, "zebra": zebra,
                   "ink": "1A202C", "gray": "718096", "line": mix(primary, "FFFFFF", 0.8), "sub_tint": mix(primary, "FFFFFF", 0.6),
                   "chart": chart},
        "fonts": {"zh_body": fonts[0], "zh_heading": fonts[1], "en": fonts[2]},
        "banner": banner, "table": table, "cover": cover, "mono": False,
    }


def uiux_rows(data_dir, query):
    path = os.path.join(data_dir, "colors.csv")
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    q = query.lower()
    return [r for r in rows if q in r["Product Type"].lower() or q in (r.get("Notes") or "").lower()]


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--id")
    ap.add_argument("--name")
    ap.add_argument("--primary")
    ap.add_argument("--accent")
    ap.add_argument("--secondary")
    ap.add_argument("--tint")
    ap.add_argument("--banner", default="filled", choices=["filled", "rule", "sidebar", "plain"])
    ap.add_argument("--table", default="zebra", choices=["zebra", "lines", "minimal"])
    ap.add_argument("--cover", default="left", choices=["center", "left", "band"])
    ap.add_argument("--from-uiux", dest="uiux")
    ap.add_argument("--uiux-data", default=os.path.join(os.path.expanduser("~"), ".claude", "skills", "ui-ux-pro-max", "data"))
    ap.add_argument("--pick", type=int, default=0, help="符合的配色列中取第幾個(從 0 起)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--out", default=OUT_DIR)
    a = ap.parse_args()

    if a.uiux:
        if not os.path.exists(os.path.join(a.uiux_data, "colors.csv")):
            sys.exit(f"找不到 ui-ux-pro-max 的 colors.csv:{a.uiux_data}(沒有該 skill 時請改用 --primary 與 --accent)")
        rows = uiux_rows(a.uiux_data, a.uiux)
        if a.list or not rows:
            for i, r in enumerate(rows):
                print(i, r["Product Type"], r["Primary"], r["Secondary"], r["Accent"], "|", (r.get("Notes") or "")[:60])
            return 0 if rows else 1
        r = rows[min(a.pick, len(rows) - 1)]
        primary, accent, secondary = r["Primary"], r["Accent"], r["Secondary"]
        inspiration = f"ui-ux-pro-max 配色:{r['Product Type']}"
    else:
        if not (a.primary and a.accent):
            ap.error("需要 --primary 與 --accent,或使用 --from-uiux")
        primary, accent, secondary, inspiration = a.primary, a.accent, a.secondary, "自訂色碼"
    if not (a.id and a.name):
        ap.error("--id 與 --name 必填")
    pack = build_pack(a.id, a.name, primary, accent, secondary, a.tint, banner=a.banner, table=a.table, cover=a.cover,
                      inspiration=inspiration)
    os.makedirs(a.out, exist_ok=True)
    path = os.path.join(a.out, a.id + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(pack, f, ensure_ascii=False, indent=2)
    print("已產生風格包:", path, "| 主色", pack["colors"]["primary"], "對比", round(contrast_with_white(pack["colors"]["primary"]), 1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
