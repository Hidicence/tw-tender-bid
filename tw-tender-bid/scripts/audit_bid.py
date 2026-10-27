# -*- coding: utf-8 -*-
"""
投標文件自動審查(交件前)。只做機器能可靠判斷的檢查,其餘請對照 references/audit-checklist.md 人工複核。

用法:
    python -I audit_bid.py <成品.docx> [--price 1062600 --price 1012000] [--project-name 案名]
        [--case-no 案號] [--items "評審項目1,評審項目2"] [--rfp <拆標文字資料夾>]
        [--blind] [--allow-terms 字樣1,字樣2] [--stale-terms 字樣3] [--out audit.json]

檢查項目(severity: blocking 必改 / warning 建議改):
  placeholder   殘留【待補】、TODO、範例字樣           blocking
  emoji         含 Emoji                                blocking
  price         金額與 --price 不一致、疑似舊價          blocking / warning
  blind         盲評案洩漏公司名、統編、負責人、Logo 字樣 blocking
  stale_case    殘留其他案件名稱(複製貼上痕跡)         warning
  names         專案名稱、案號缺漏;公司名稱寫法不一致   warning
  coverage      評審項目沒有對應標題或內文               warning
  mainland      簡體字或大陸用語                        warning
  format        對照 --format 的招標格式規格(紙張、邊界、字體、字級、行距、頁碼、黑白、頁數) blocking

結束碼:有 blocking 時為 1,否則 0。
"""
import argparse, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract_rfp import read_docx  # noqa: E402
import profile_loader  # noqa: E402

# 只抓彩色 Emoji;單色幾何與符號字元(◆ ▸ ▪ ◎ ★ ✓ 等)是使用者允許的,不算
EMOJI = re.compile(
    "[\U0001F000-\U0001FAFF"
    "✅❌❎⭐⭕⚠❗❓❔❕✨⚡⬛⬜"
    "⏩-⏳⌚⌛☔☕⚓⚪⚫⚽⚾⛄⛅⛔"
    "⛪⛲⛳⛵⛺⛽✊✋➕-➗➰➿❤]"
    "|.️"
)
PLACEHOLDER = re.compile(r"【待補[^】]*】|待補充|TODO|XXX|範例文字|示意用|Lorem ipsum|○○|△△|□□")
# 簡體專有字:這些字形繁體文件不會出現(皆為簡化後的字形)
SIMPLIFIED_ONLY = set("这们实发开关与为业务软视频质认库户络员对时说话请见观览设计线统项")
MAINLAND_TERMS = ["軟件", "信息", "視頻", "質量", "默認", "數據庫", "用戶", "激活", "屏幕", "網絡", "程序員",
                  "打印", "硬盤", "鼠標", "文檔", "支持(", "優化方案", "解決方案提供商", "項目經理", "接口",
                  "渠道", "落地", "賦能", "抓手", "閉環"]
MONEY = re.compile(r"(?<![\d,])(\d{1,3}(?:,\d{3})+|\d{5,})(?![\d,])")
CO = re.compile(r"[一-鿿A-Za-z0-9（）()·]{2,20}(?:股份有限公司|有限公司)")


def full_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return read_docx(path)[0]
    return open(path, encoding="utf-8", errors="replace").read()


def add(findings, sev, code, msg, where=""):
    findings.append({"severity": sev, "check": code, "message": msg, "where": where})


def ctx(text, m, w=24):
    s = max(0, m.start() - w)
    return text[s:m.end() + w].replace("\n", " ")


def check_placeholders(t, F):
    ms = list(PLACEHOLDER.finditer(t))
    for m in ms[:50]:
        add(F, "blocking", "placeholder", f"殘留待補或範例字樣:{m.group(0)}", ctx(t, m))
    if len(ms) > 50:
        add(F, "blocking", "placeholder", f"另有 {len(ms) - 50} 處未列出")


def check_emoji(t, F):
    seen = set()
    for m in EMOJI.finditer(t):
        if m.group(0) in seen:
            continue
        seen.add(m.group(0))
        add(F, "blocking", "emoji", f"含 Emoji:{m.group(0)}(U+{ord(m.group(0)):04X})", ctx(t, m))


def _is_id(token, line, pos):
    """沒有千分位逗號、前面沒有貨幣標記的數字(統一編號、電話、案號)不當金額。"""
    if "," in token:
        return False
    before = line[max(0, pos - 6):pos]
    return not re.search(r"新台幣|NT\$|\$|＄|元", before)


def check_prices(t, prices, F):
    if not prices:
        return
    # 強訊號:明講總價、總報價、契約金額、投標金額的行 -> 金額不符就是必改
    # 弱訊號:只提到含稅、合計、報價的行,可能是分項金額 -> 提醒人工確認
    # 強訊號句型:「總價/總報價/契約金額/投標金額/總計」後面 20 字內(不跨數字)接「為、:、新台幣、NT$」再接金額
    strong_re = re.compile(r"(?:專案總報價|總報價|總價|契約金額|投標金額|總計)[^\d。;\n]{0,20}?(?:為|[:：]|新台幣|NT\$|＄|\$)\s*(?:新台幣|NT\$|\$)?\s*(\d{1,3}(?:,\d{3})+|\d{5,})")
    weak_re = re.compile(r"含稅|合計|報價")
    strong, weak = {}, {}
    for line in t.split("\n"):
        for m in strong_re.finditer(line):
            v = int(m.group(1).replace(",", ""))
            if v >= 10000:
                strong.setdefault(v, []).append(line.strip()[:60])
        if weak_re.search(line):
            for m in MONEY.finditer(line):
                v = int(m.group(1).replace(",", ""))
                if v >= 10000 and not _is_id(m.group(1), line, m.start()):
                    weak.setdefault(v, []).append(line.strip()[:60])
    allowed = set(prices)
    present = {int(m.group(1).replace(",", "")) for m in MONEY.finditer(t)}
    for p in prices:
        if p not in present:
            add(F, "blocking", "price", f"最終金額 {p:,} 在文件中完全沒出現")
    for v, lines in strong.items():
        if v not in allowed:
            add(F, "blocking", "price",
                f"總價相關行出現未核准的金額 {v:,}(疑似舊價或誤植)", " / ".join(lines[:2]))
    for v, lines in weak.items():
        if v not in allowed and v not in strong:
            add(F, "warning", "price",
                f"含稅或報價行出現金額 {v:,},若為分項金額請對照標價單確認", " / ".join(lines[:2]))


def check_blind(t, company, F):
    terms = [company.get(k, "") for k in ("name", "tax_id", "owner", "name_en", "phone", "email")]
    terms += company.get("brand_terms", [])
    for term in [x for x in terms if x and not x.startswith("【待補")]:
        for m in re.finditer(re.escape(term), t, flags=re.I):
            add(F, "blocking", "blind", f"盲評案出現可辨識廠商資訊:{term}", ctx(t, m))
            break


def check_stale(t, terms, allow, F):
    """terms 來自 profile 的 past_case_keywords(公司過往案件的關鍵字)與 --stale-terms。"""
    for term in terms:
        if term in allow:
            continue
        m = re.search(re.escape(term), t)
        if m:
            add(F, "warning", "stale_case", f"出現其他案件字樣「{term}」,確認不是複製貼上殘留", ctx(t, m))


def check_names(t, project, case_no, company, F):
    if project and project not in t:
        add(F, "warning", "names", f"文件中找不到完整專案名稱:{project}")
    if case_no and case_no not in t:
        add(F, "warning", "names", f"文件中找不到案號:{case_no}")
    # 以公司名前兩字開頭、以「公司」結尾的字串,必須等於 profile 的全稱(縮寫或漏字為不一致寫法)
    variants = []
    if company["name"] and not company["name"].startswith("【待補"):
        stem = re.escape(company["name"][:2])
        variants = sorted({m.group(0) for m in re.finditer(stem + r"[一-鿿]{0,6}公司", t)} - {company["name"]})
    if variants:
        add(F, "warning", "names", f"公司名稱寫法不一致:{variants}(標準寫法 {company['name']})")


PAPERS_CM = {"A4": (21.0, 29.7), "A3": (29.7, 42.0), "B4": (25.7, 36.4), "B5": (18.2, 25.7), "Letter": (21.59, 27.94)}


def _iter_paragraphs(doc):
    """(段落, 位置) ;位置為 body / table / header / footer。"""
    for para in doc.paragraphs:
        yield para, "body"
    for tb in doc.tables:
        for row in tb.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for para in cell.paragraphs:
                    yield para, "table"
    for sec in doc.sections:
        for para in sec.header.paragraphs:
            yield para, "header"
        for para in sec.footer.paragraphs:
            yield para, "footer"


def _run_fonts(run, doc):
    """回傳 (中文字體, 西文字體, 字級 pt)。run 沒明講時退回 Normal 樣式。"""
    from docx.oxml.ns import qn
    rpr = run._element.rPr
    ea = ascii_ = None
    if rpr is not None and rpr.rFonts is not None:
        ea = rpr.rFonts.get(qn("w:eastAsia"))
        ascii_ = rpr.rFonts.get(qn("w:ascii"))
    st = doc.styles["Normal"]
    st_rpr = st.element.rPr
    if st_rpr is not None and st_rpr.rFonts is not None:
        ea = ea or st_rpr.rFonts.get(qn("w:eastAsia"))
        ascii_ = ascii_ or st_rpr.rFonts.get(qn("w:ascii"))
    size = run.font.size.pt if run.font.size else (st.font.size.pt if st.font.size else None)
    return ea, ascii_, size


def check_format(path, fmt, pdf, F):
    """核對招標文件的格式規格(format.json)。這些是招標文件的硬性規定,不符一律 blocking。"""
    from docx import Document
    from docx.oxml.ns import qn
    d = Document(path)
    where = os.path.basename(path)
    sec = d.sections[0]

    # 紙張
    want = fmt.get("paper")
    if want in PAPERS_CM:
        w, h = PAPERS_CM[want]
        if fmt.get("orientation") == "landscape":
            w, h = h, w
        gw, gh = sec.page_width.cm, sec.page_height.cm
        if abs(gw - w) > 0.15 or abs(gh - h) > 0.15:
            add(F, "blocking", "format", f"紙張不符:規格 {want} {w}x{h} cm,實際 {gw:.1f}x{gh:.1f} cm", where)

    # 邊界
    m = fmt.get("margins_cm") or {}
    for key, val in (("top", sec.top_margin), ("bottom", sec.bottom_margin), ("left", sec.left_margin), ("right", sec.right_margin)):
        if key in m and abs(val.cm - m[key]) > 0.1:
            add(F, "blocking", "format", f"邊界 {key} 不符:規格 {m[key]} cm,實際 {val.cm:.2f} cm", where)
    if "gutter_cm" in fmt and abs(sec.gutter.cm - fmt["gutter_cm"]) > 0.1:
        add(F, "blocking", "format", f"裝訂邊不符:規格 {fmt['gutter_cm']} cm,實際 {sec.gutter.cm:.2f} cm", where)

    font = fmt.get("font") or {}
    want_zh, want_en = font.get("zh"), font.get("en")
    body_pt, table_pt = font.get("size_pt"), font.get("table_pt", font.get("size_pt"))

    total = bad_zh = bad_en = small = colored = 0
    zh_ex, en_ex, small_ex = [], [], []
    sizes = {}
    for para, loc in _iter_paragraphs(d):
        for r in para.runs:
            txt = r.text.strip()
            if not txt:
                continue
            ea, asc, size = _run_fonts(r, d)
            n = len(txt)
            if loc in ("header", "footer"):
                continue  # 頁首頁尾字級常較小,只核對字體
            total += n
            if size:
                sizes[size] = sizes.get(size, 0) + n
            has_cjk = any("\u4e00" <= c <= "\u9fff" for c in txt)
            if want_zh and has_cjk and ea != want_zh:
                bad_zh += n
                if len(zh_ex) < 3:
                    zh_ex.append(f"{ea}:{txt[:12]}")
            if want_en and asc and any(c.isascii() and c.isalnum() for c in txt) and asc != want_en:
                bad_en += n
                if len(en_ex) < 3:
                    en_ex.append(f"{asc}:{txt[:12]}")
            need = table_pt if loc == "table" else body_pt
            if need and size and size + 0.01 < need and not r.font.superscript:
                small += n
                if len(small_ex) < 3:
                    small_ex.append(f"{size}pt:{txt[:12]}")
            if fmt.get("color_mode") == "mono" and r.font.color is not None and r.font.color.rgb is not None:
                c = r.font.color.rgb
                if not (c[0] == c[1] == c[2]):
                    colored += n
    if want_zh and bad_zh:
        add(F, "blocking", "format", f"中文字體不符規格「{want_zh}」:{bad_zh} 字使用其他字體", "; ".join(zh_ex))
    if want_en and bad_en:
        add(F, "blocking", "format", f"西文與數字字體不符規格「{want_en}」:{bad_en} 字", "; ".join(en_ex))
    if small:
        add(F, "blocking", "format", f"{small} 字小於規格字級(內文 {body_pt}、表格 {table_pt} pt)", "; ".join(small_ex))
    if body_pt and sizes:
        dom = max(sizes, key=sizes.get)
        if abs(dom - body_pt) > 0.01 and abs(dom - (table_pt or body_pt)) > 0.01:
            add(F, "warning", "format", f"最多字數的字級是 {dom} pt,與規格內文 {body_pt} pt 不同,請確認")
    if fmt.get("color_mode") == "mono" and colored:
        add(F, "blocking", "format", f"規格要求黑白(單色),但有 {colored} 字使用彩色文字")

    # 行距(抽查內文段落)
    want_ls = fmt.get("line_spacing")
    want_exact = fmt.get("line_exact_pt")
    if want_ls or want_exact:
        mism = 0
        n_checked = 0
        for para in d.paragraphs:
            if not para.text.strip() or len(para.text) < 30:
                continue
            sp = para._p.pPr.find(qn("w:spacing")) if para._p.pPr is not None else None
            if sp is None:
                continue
            line, rule = sp.get(qn("w:line")), sp.get(qn("w:lineRule"))
            if not line:
                continue
            n_checked += 1
            if want_exact:
                ok = (rule == "exact" and abs(int(line) - want_exact * 20) <= 10)
            else:
                ok = (rule in (None, "auto") and abs(int(line) / 240 - want_ls) <= 0.05)
            mism += 0 if ok else 1
        if n_checked and mism:
            add(F, "blocking", "format", f"行距不符規格:抽查 {n_checked} 個內文段落中有 {mism} 個不符", where)

    # 頁碼
    if fmt.get("page_numbers"):
        xml = "".join(sec.footer._element.xml for sec in d.sections) + "".join(sec.header._element.xml for sec in d.sections)
        if "PAGE" not in xml:
            add(F, "blocking", "format", "規格要求頁碼,但頁首頁尾找不到頁碼欄位", where)

    # 頁數上限(需要 PDF)
    if fmt.get("page_limit"):
        if pdf and os.path.exists(pdf):
            import pdfplumber
            with pdfplumber.open(pdf) as pp:
                n = len(pp.pages)
            excl = int(fmt.get("page_limit_exclude_pages", 0))
            if n - excl > fmt["page_limit"]:
                add(F, "blocking", "format",
                    f"頁數 {n}(扣除封面與目錄等不計頁數 {excl})超過規格上限 {fmt['page_limit']};附件是否計入依須知認定",
                    os.path.basename(pdf))
        else:
            add(F, "warning", "format", f"規格有頁數上限 {fmt['page_limit']},但未提供 --pdf,無法核對頁數")


def check_items(t, items, F):
    flat = re.sub(r"\s+", "", t)
    for it in items:
        key = re.sub(r"\s+", "", it)
        probe = key[:8] if len(key) > 8 else key
        if probe not in flat:
            add(F, "warning", "coverage", f"找不到評審項目「{it}」對應的標題或內文")


def check_mainland(t, F):
    chars = sorted({c for c in t if c in SIMPLIFIED_ONLY})
    if chars:
        add(F, "warning", "mainland", f"疑似簡體字:{''.join(chars)}")
    for term in MAINLAND_TERMS:
        m = re.search(re.escape(term), t)
        if m:
            add(F, "warning", "mainland", f"大陸用語「{term}」,改為台灣用語", ctx(t, m))


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--price", type=int, action="append", default=[])
    ap.add_argument("--project-name")
    ap.add_argument("--case-no")
    ap.add_argument("--items")
    ap.add_argument("--rfp")
    ap.add_argument("--blind", action="store_true")
    ap.add_argument("--allow-terms", default="", help="放行的舊案字樣(本案相關者)")
    ap.add_argument("--stale-terms", default="", help="額外要檢查的舊案字樣,逗號分隔")
    ap.add_argument("--format", help="招標文件的格式規格 format.json,核對紙張、邊界、字體、字級、行距、頁碼、黑白")
    ap.add_argument("--pdf", help="成品 PDF,用來核對頁數上限")
    ap.add_argument("--out")
    a = ap.parse_args()

    company = profile_loader.load()
    if company["_is_template"]:
        print("注意:未設定公司 profile,使用空白模板(見 profile_loader.py)。公司相關檢查會不完整。")
    t = full_text(a.doc)
    F = []
    check_placeholders(t, F)
    check_emoji(t, F)
    check_prices(t, a.price, F)
    if a.blind:
        check_blind(t, company, F)
    stale_terms = list(company.get("past_case_keywords", [])) + [x for x in a.stale_terms.split(",") if x]
    check_stale(t, stale_terms, set(x for x in a.allow_terms.split(",") if x), F)
    check_names(t, a.project_name, a.case_no, company, F)
    if a.items:
        check_items(t, [x.strip() for x in a.items.split(",") if x.strip()], F)
    check_mainland(t, F)
    if a.format:
        if a.doc.lower().endswith(".docx"):
            with open(a.format, encoding="utf-8") as fh:
                check_format(a.doc, json.load(fh), a.pdf, F)
        else:
            add(F, "warning", "format", "格式規格核對只支援 .docx")

    blocking = [f for f in F if f["severity"] == "blocking"]
    warning = [f for f in F if f["severity"] == "warning"]
    report = {"doc": a.doc, "chars": len(t), "blocking": len(blocking), "warning": len(warning), "findings": F}
    if a.out:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        json.dump(report, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"審查:{a.doc}  字數 {len(t)}  blocking {len(blocking)}  warning {len(warning)}")
    for f in blocking + warning:
        print(f"[{f['severity']}] {f['check']}: {f['message']}" + (f"  <- {f['where']}" if f["where"] else ""))
    if not F:
        print("自動檢查未發現問題(仍須依 audit-checklist.md 人工複核)")
    return 1 if blocking else 0


if __name__ == "__main__":
    sys.exit(main())
