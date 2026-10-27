# -*- coding: utf-8 -*-
"""
bid_kit.py -- 投標文件共用版面引擎(Word / matplotlib)

特色:章次橫幅、次標題、表格(自動辨識分類列與小計列)、重點框、封面、頁首頁尾與頁碼、
簽章區、圖表嵌入。公司資料由 profile_loader 讀取(換公司只要換 profile)。

兩層設計:
  規格層  來自招標文件(紙張、邊界、字體、各級字級、行距、黑白或彩色、頁首頁碼),
          用 configure(format={...}) 或 apply_format() 設定,優先於風格。
  風格層  assets/styles/*.json 風格包(配色、字體搭配、標題樣式、表格樣式、封面版式),
          用 configure(style="swiss-grid") 設定;list_styles() 列出可用風格。

用法:
    import bid_kit as K
    K.configure(case_title="...", case_short="...", client="...", subtitle="...",
                cover_date="中華民國 115 年 8 月", sign_date="", partner=None,
                blind=False, style="navy-institutional", format=None, cover_image=None)
    doc = K.new_doc()
    K.build_cover(doc, "服務建議書")
    K.banner(doc, "壹、專案摘要"); K.body(doc, "...")
    K.add_header_footer(doc); K.save_doc(doc, "out.docx")

規則:不用彩色 Emoji、不用漸層與彩色圖示;簽名與印章只在呼叫 sign_block(seals=True) 或
add_seals() 時才插入,且須經使用者逐份確認(見 SKILL.md 鐵則)。
"""

import os, io, sys, json
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import profile_loader  # noqa: E402

# ── 公司資料(來自 profile)────────────────────────────────────────
PROFILE = profile_loader.load()
BIDDER = {k: PROFILE.get(k, "") for k in ("name", "name_en", "tax_id", "owner", "address", "phone", "email", "founded")}
_SS = PROFILE.get("seals_and_signature", {})
SIGNATURE    = _SS.get("signature_owner", "")
SEAL_COMPANY = _SS.get("seal_company", "")
SEAL_OWNER   = _SS.get("seal_owner", "")

# ── 案件資訊(由 configure() 設定)────────────────────────────────
CASE_TITLE = ""
CASE_SHORT = ""
CLIENT = ""
CASE_SUBTITLE = ""
PARTNER = None        # 例:{"name": "協力單位全名"}
COVER_DATE = ""
SIGN_DATE = ""
BLIND = False         # 盲評案:封面與簽章區不顯示公司名
COVER_IMAGE = None    # 封面主視覺圖路徑(選用)


def _rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _gray_hex(h):
    """轉成同亮度的灰階(黑白列印與盲評用)。"""
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    y = int(0.299 * r + 0.587 * g + 0.114 * b)
    return "%02X%02X%02X" % (y, y, y)


STYLES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "styles")
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ── 規格層:來自招標文件,優先於風格(由 apply_format() 設定)─────────────
FORMAT = {}
PAPER = "A4"
LANDSCAPE = False
MARGINS = dict(top=2.2, bottom=2.0, left=2.2, right=2.2)
GUTTER_CM = 0.0
BODY_PT = 11.0
SZ = {}                # 各級字級,見 _recompute_sizes()
LINE_TW = 360          # 行距數值:multiple 時 240 = 單行距;exact 時為 pt * 20
LINE_RULE = "auto"     # auto(倍數)或 exact(固定行高)
SHOW_HEADER = True
SHOW_PAGENUM = True
MONO_FORCED = False
FONT_ZH_OVERRIDE = None
FONT_ZH_HEAD_OVERRIDE = None
FIRST_INDENT = 0.0     # 首行縮排字元數(規格層 indent_chars)
HEADER_CHAPTER = True  # 頁首右側顯示目前章名(STYLEREF)
FONT_EN_OVERRIDE = None
PAPERS = {"A4": (21.0, 29.7), "A3": (29.7, 42.0), "B4": (25.7, 36.4), "B5": (18.2, 25.7), "Letter": (21.59, 27.94)}
# matplotlib 用的 Windows 字型檔(圖表裡的字也要符合字體規定)
_MPL_FONT_FILES = {"標楷體": "kaiu.ttf", "新細明體": "mingliu.ttc", "細明體": "mingliu.ttc",
                   "微軟正黑體": "msjh.ttc", "Microsoft JhengHei": "msjh.ttc"}

# ── 風格層:可替換(由 set_style() 設定)───────────────────────────────
ZH = "微軟正黑體"
ZH_HEAD = "微軟正黑體"
EN = "Calibri"
BANNER_STYLE = "filled"   # filled | rule | sidebar | plain
TABLE_STYLE = "zebra"     # zebra | lines | minimal
COVER_LAYOUT = "center"   # center | left | band
COVER_ART = "none"       # none | blocks | rings | bars | waves | dots(用色票自動生成的幾何圖形)
_CURRENT_STYLE = None
_ALIASES = {"navy": "navy-institutional", "green": "green-civic"}


def _recompute_sizes():
    """各級字級。規格層可個別指定,沒指定的由內文字級推算。"""
    global SZ
    f = (FORMAT.get("font") or {})
    SZ = {
        "body": BODY_PT,
        "table": float(f.get("table_pt", BODY_PT - 1)),
        "h1": float(f.get("h1_pt", BODY_PT + 5)),
        "h2": float(f.get("h2_pt", BODY_PT + 2)),
        "h3": float(f.get("h3_pt", BODY_PT + 0.5)),
        "caption": float(f.get("caption_pt", max(8.0, BODY_PT - 2))),
        "cover_title": float(f.get("cover_title_pt", BODY_PT + 15)),
    }


def list_styles():
    out = []
    if os.path.isdir(STYLES_DIR):
        for f in sorted(os.listdir(STYLES_DIR)):
            if f.endswith(".json"):
                with open(os.path.join(STYLES_DIR, f), encoding="utf-8") as fh:
                    d = json.load(fh)
                out.append((d["id"], d["name"], d.get("school", ""), d.get("best_for", "")))
    return out


def load_style(x):
    """x 可以是風格 id、舊的主題別名(navy/green)、風格包 dict,或風格包 JSON 的路徑。"""
    if isinstance(x, dict):
        return x
    x = _ALIASES.get(x, x)
    path = x if os.path.isfile(str(x)) else os.path.join(STYLES_DIR, f"{x}.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def set_style(x):
    """套用風格包。所有版面函式在呼叫時才讀這些全域值,切換立即生效。規格層(FORMAT)的覆寫優先。"""
    global NAVY, STEEL, TEAL, PALE, MIST, INK, GRAY, LINE, ACCENT, SUB_TINT, CHART_COLORS
    global C_NAVY, C_STEEL, C_TEAL, C_PALE, C_GRAY, C_ACC
    global ZH, ZH_HEAD, EN, BANNER_STYLE, TABLE_STYLE, COVER_LAYOUT, COVER_ART, _CURRENT_STYLE
    pack = load_style(x)
    _CURRENT_STYLE = pack
    c = dict(pack["colors"])
    chart = list(c["chart"])
    if MONO_FORCED or pack.get("mono"):
        c = {k: (_gray_hex(v) if isinstance(v, str) else v) for k, v in c.items()}
        chart = [("#" + _gray_hex(v)) for v in chart]
    NAVY, STEEL, TEAL = _rgb(c["primary"]), _rgb(c["secondary"]), _rgb(c["tertiary"])
    ACCENT, PALE, MIST = _rgb(c["accent"]), _rgb(c["tint"]), _rgb(c["zebra"])
    INK, GRAY, LINE, SUB_TINT = _rgb(c["ink"]), _rgb(c["gray"]), _rgb(c["line"]), _rgb(c["sub_tint"])
    CHART_COLORS = chart
    C_NAVY, C_STEEL, C_TEAL = "#" + c["primary"], "#" + c["secondary"], "#" + c["tertiary"]
    C_PALE, C_ACC, C_GRAY = "#" + c["tint"], "#" + c["accent"], "#" + c["gray"]
    f = pack.get("fonts", {})
    EN = FONT_EN_OVERRIDE or f.get("en", "Calibri")
    ZH = FONT_ZH_OVERRIDE or f.get("zh_body", "微軟正黑體")
    ZH_HEAD = FONT_ZH_HEAD_OVERRIDE or FONT_ZH_OVERRIDE or f.get("zh_heading", ZH)
    BANNER_STYLE = pack.get("banner", "filled")
    TABLE_STYLE = pack.get("table", "zebra")
    COVER_LAYOUT = pack.get("cover", "center")
    COVER_ART = pack.get("art", "none")
    _recompute_sizes()
    chart_font()


def chart_font():
    """讓圖表(matplotlib)的字體跟文件一致;找不到字型檔就退回微軟正黑體。"""
    from matplotlib import font_manager as fm
    names = []
    fp = os.path.join(r"C:\Windows\Fonts", _MPL_FONT_FILES.get(ZH, ""))
    if _MPL_FONT_FILES.get(ZH) and os.path.exists(fp):
        try:
            fm.fontManager.addfont(fp)
            names.append(fm.FontProperties(fname=fp).get_name())
        except Exception:
            pass
    plt.rcParams["font.family"] = names + ["Microsoft JhengHei", "Microsoft YaHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


def apply_format(fmt):
    """套用招標文件的格式規格(每一項都應在 format.json 註明須知出處,見 references/style-system.md):
        {"paper": "A4", "orientation": "portrait",
         "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5}, "gutter_cm": 0.5,
         "font": {"zh": "標楷體", "zh_heading": "標楷體", "en": "Times New Roman",
                  "size_pt": 12, "h1_pt": 16, "h2_pt": 14, "h3_pt": 12, "table_pt": 12, "caption_pt": 10},
         "line_spacing": 1.5,              倍數;或改用 "line_exact_pt": 20 固定行高
         "color_mode": "mono", "header": false, "page_numbers": true}
    未提供的項目維持預設。規格層優先於風格層,風格包的字體與字級會被覆寫。"""
    global FORMAT, PAPER, LANDSCAPE, MARGINS, GUTTER_CM, BODY_PT, LINE_TW, LINE_RULE
    global SHOW_HEADER, SHOW_PAGENUM, MONO_FORCED, FONT_ZH_OVERRIDE, FONT_ZH_HEAD_OVERRIDE, FONT_EN_OVERRIDE, FIRST_INDENT
    FORMAT = fmt or {}
    PAPER = FORMAT.get("paper", "A4")
    LANDSCAPE = (FORMAT.get("orientation") == "landscape")
    m = FORMAT.get("margins_cm") or {}
    MARGINS = dict(top=m.get("top", 2.2), bottom=m.get("bottom", 2.0), left=m.get("left", 2.2), right=m.get("right", 2.2))
    GUTTER_CM = float(FORMAT.get("gutter_cm", 0.0))
    font = FORMAT.get("font") or {}
    BODY_PT = float(font.get("size_pt", 11.0))
    FONT_ZH_OVERRIDE = font.get("zh") or None
    FONT_ZH_HEAD_OVERRIDE = font.get("zh_heading") or None
    FONT_EN_OVERRIDE = font.get("en") or None
    if FORMAT.get("line_exact_pt"):
        LINE_RULE, LINE_TW = "exact", int(round(float(FORMAT["line_exact_pt"]) * 20))
    else:
        LINE_RULE, LINE_TW = "auto", int(round(float(FORMAT.get("line_spacing", 1.5)) * 240))
    SHOW_HEADER = bool(FORMAT.get("header", True))
    FIRST_INDENT = float(FORMAT.get("indent_chars", 0))
    SHOW_PAGENUM = bool(FORMAT.get("page_numbers", True))
    MONO_FORCED = (FORMAT.get("color_mode") == "mono")
    if _CURRENT_STYLE is not None:
        set_style(_CURRENT_STYLE)
    else:
        _recompute_sizes()


set_style(PROFILE.get("default_theme", "navy-institutional"))


def configure(case_title="", case_short="", client="", subtitle="",
              cover_date="", sign_date="", partner=None, blind=False,
              style=None, theme=None, format=None, cover_image=None):
    """style:風格 id 或風格包;theme 為舊參數別名。format:招標文件的格式規格(見 apply_format)。
    先設 format 再設 style,確保規格覆寫風格。"""
    global CASE_TITLE, CASE_SHORT, CLIENT, CASE_SUBTITLE, PARTNER, COVER_DATE, SIGN_DATE, BLIND, COVER_IMAGE
    CASE_TITLE, CASE_SHORT, CLIENT, CASE_SUBTITLE = case_title, case_short, client, subtitle
    PARTNER, COVER_DATE, SIGN_DATE, BLIND = partner, cover_date, sign_date, blind
    COVER_IMAGE = cover_image
    if format is not None:
        apply_format(format)
    if style or theme:
        set_style(style or theme)


def _setup_mpl():
    from matplotlib import font_manager as fm
    for fn in ["msjh.ttc", "msjhbd.ttc", "mingliu.ttc"]:
        fp = os.path.join(r"C:\Windows\Fonts", fn)
        if os.path.exists(fp):
            try: fm.fontManager.addfont(fp)
            except Exception: pass
    plt.rcParams["font.family"] = ["Microsoft JhengHei", "Microsoft YaHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


_setup_mpl()



def text_width_cm():
    w, h = PAPERS.get(PAPER, PAPERS["A4"])
    if LANDSCAPE:
        w = h
    return w - MARGINS["left"] - MARGINS["right"] - GUTTER_CM


def _border_el(tag, spec=None):
    el = OxmlElement(f"w:{tag}")
    if spec:
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(spec[0]))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), str(spec[1]))
    else:
        el.set(qn("w:val"), "nil")
    return el


def _insert_tblpr(t, el):
    """依 schema 順序,把 tblBorders、tblCellMar 放在 tblLook 之前。"""
    tblPr = t._tbl.tblPr
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(el)
    else:
        tblPr.append(el)


def set_table_borders(t, top=None, bottom=None, insideH=None, insideV=None, left=None, right=None, pad=(50, 110)):
    """spec 為 (sz, 色碼);None 表示無線。sz 單位為 1/8 pt(4 = 0.5pt)。pad = (上下, 左右) 儲存格邊距(twips)。"""
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")) + tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    b = OxmlElement("w:tblBorders")
    for tag, spec in (("top", top), ("left", left), ("bottom", bottom), ("right", right), ("insideH", insideH), ("insideV", insideV)):
        b.append(_border_el(tag, spec))
    _insert_tblpr(t, b)
    m = OxmlElement("w:tblCellMar")
    for tag, v in (("top", pad[0]), ("left", pad[1]), ("bottom", pad[0]), ("right", pad[1])):
        e = OxmlElement(f"w:{tag}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    _insert_tblpr(t, m)


def cell_border(cell, **sides):
    """儲存格個別邊框,例:cell_border(c, bottom=(12, '1A365D'), top=None)。放在 shd 之前以符合 schema 順序。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    b = OxmlElement("w:tcBorders")
    for tag in ("top", "left", "bottom", "right"):
        if tag in sides:
            b.append(_border_el(tag, sides[tag]))
    shd = tcPr.find(qn("w:shd"))
    if shd is not None:
        shd.addprevious(b)
    else:
        tcPr.append(b)


def table_rows_flow(t, header_rows=1):
    """表頭跨頁重複、每列不被拆到兩頁。"""
    for i, row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if i < header_rows:
            trPr.append(OxmlElement("w:tblHeader"))


def style_table(t, header_rows=1):
    """依風格包的 table 樣式設定邊框:
       zebra   無直線,細橫線,底線加粗      lines  上下粗線,表頭下粗線,細橫線      minimal  三線表"""
    prim, ink, line = str(NAVY), str(INK), str(LINE)
    if TABLE_STYLE == "minimal":
        set_table_borders(t, top=(12, ink), bottom=(12, ink))
        for c in t.rows[0].cells:
            cell_border(c, bottom=(6, ink))
    elif TABLE_STYLE == "lines":
        set_table_borders(t, top=(12, prim), bottom=(12, prim), insideH=(4, line))
        for c in t.rows[0].cells:
            cell_border(c, bottom=(12, prim))
    else:
        set_table_borders(t, top=(4, line), bottom=(12, prim), insideH=(4, line))
    table_rows_flow(t, header_rows)


def _complex_field(para, instr, cached="", size=None, color=None, bold=False):
    """複雜欄位(begin / instr / separate / 快取文字 / end),快取文字讓未更新欄位時也看得到內容。"""
    def mk(kind):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        r.append(f)
        return r
    r_instr = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r_instr.append(it)
    para._p.append(mk("begin"))
    para._p.append(r_instr)
    para._p.append(mk("separate"))
    ph = para.add_run(cached)
    ph.font.size = Pt(size if size is not None else SZ["body"])
    ph.bold = bold
    if color:
        ph.font.color.rgb = color
    ph.font.name = EN
    ph._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    para._p.append(mk("end"))


_SEQ = {}


def _set_style_font(style, size, bold=False, color=None, zh=None, italic=False):
    f = style.font
    f.name = EN
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    if color is not None:
        f.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for a in list(rf.attrib):
        if "Theme" in a or "theme" in a:  # 去掉內建標題的主題字型,否則會蓋過我們指定的字體
            del rf.attrib[a]
    rf.set(qn("w:ascii"), EN)
    rf.set(qn("w:hAnsi"), EN)
    rf.set(qn("w:eastAsia"), zh or ZH)
    col = rpr.find(qn("w:color"))
    if col is not None:
        for a in ("themeColor", "themeShade", "themeTint"):
            k = qn("w:" + a)
            if k in col.attrib:
                del col.attrib[k]


def define_styles(doc):
    """把版面設定寫進 Word 樣式(Normal、標題 1 到 3、圖說),而不是只靠逐段的直接格式。
    這樣 Word 的導覽窗格、目錄、PDF 書籤都能用,使用者之後改字體也只需改樣式。"""
    S = doc.styles
    _set_style_font(S["Normal"], SZ["body"], color=INK)
    _set_style_font(S["Heading 1"], SZ["h1"], bold=True, color=NAVY, zh=ZH_HEAD)
    _set_style_font(S["Heading 2"], SZ["h2"], bold=True, color=STEEL, zh=ZH_HEAD)
    _set_style_font(S["Heading 3"], SZ["h3"], bold=True, color=TEAL, zh=ZH_HEAD)
    _set_style_font(S["Caption"], SZ["caption"], color=GRAY)


# ════════════════════════════════════════════════════════════════
#  Word 低階輔助
# ════════════════════════════════════════════════════════════════
def set_cell_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "%02X%02X%02X" % (rgb[0], rgb[1], rgb[2]))
    tcPr.append(shd)


def spc(para, before=0, after=0, lr=None, ln=None):
    pPr = para._p.get_or_add_pPr()
    s = OxmlElement("w:spacing")
    s.set(qn("w:before"), str(before))
    s.set(qn("w:after"), str(after))
    if lr and ln:
        s.set(qn("w:lineRule"), lr)
        s.set(qn("w:line"), str(ln))
    pPr.append(s)


def run(para, text, bold=False, italic=False, size=None, color=None, font=None):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size if size is not None else SZ["body"])
    if color:
        r.font.color.rgb = color
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font or ZH)
    return r


def para_border(para, side, sz, color, space=4):
    pPr = para._p.get_or_add_pPr()
    pb = pPr.find(qn("w:pBdr"))
    if pb is None:
        pb = OxmlElement("w:pBdr")
        pPr.append(pb)
    el = OxmlElement("w:%s" % side)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), "%02X%02X%02X" % (color[0], color[1], color[2]))
    pb.append(el)


def para_bg(para, rgb):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "%02X%02X%02X" % (rgb[0], rgb[1], rgb[2]))
    pPr.append(shd)


def keep_with_next(para):
    pPr = para._p.get_or_add_pPr()
    el = OxmlElement("w:keepNext")
    pPr.append(el)


# ── 分頁與空白段落清理 ──────────────────────────────────────────
def _is_blank_para(el):
    """判斷段落是否為空白（無文字、無圖片、無分頁符）"""
    if el.tag != qn("w:p"):
        return False
    if "".join(el.itertext()).strip():
        return False
    for tag in ("w:drawing", "w:pict", "w:object", "w:br"):
        if el.findall(".//" + qn(tag)):
            return False
    return True


def strip_trailing_blanks(doc):
    """移除文件尾端連續的空白段落。
    tbl() 於表格後會補一個空段落作間距，若剛好落在章節結尾，
    會被推到下一頁而產生整頁空白，故於分頁前與存檔前清除。"""
    body = doc.element.body
    removed = 0
    for el in reversed(list(body)):
        if el.tag == qn("w:sectPr"):
            continue
        if _is_blank_para(el):
            body.remove(el)
            removed += 1
        else:
            break
    return removed


def page_break(doc):
    """章節分頁：先清掉尾端空白段落，避免產生整頁空白"""
    strip_trailing_blanks(doc)
    doc.add_page_break()


def save_doc(doc, path):
    """存檔前清尾端空白並開啟欄位自動更新"""
    strip_trailing_blanks(doc)
    set_update_fields_on_open(doc)
    doc.save(path)


# ── 書籤與交叉參照頁碼 ──────────────────────────────────────────
_BM_ID = [1000]


def add_bookmark(para, name):
    """在段落起始處插入書籤，供目錄 PAGEREF 參照"""
    _BM_ID[0] += 1
    bid = str(_BM_ID[0])
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bid)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bid)
    para._p.insert(0, start)
    para._p.append(end)
    return name


def add_pageref(para, bookmark, size=10, color=None, bold=False):
    """插入 PAGEREF 欄位，顯示書籤所在頁碼（Word 開啟或 F9 時自動更新）"""
    def _mk(kind):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        r.append(f)
        return r

    def _instr():
        r = OxmlElement("w:r")
        t = OxmlElement("w:instrText")
        t.set(qn("xml:space"), "preserve")
        t.text = f" PAGEREF {bookmark} \\h "
        r.append(t)
        return r

    para._p.append(_mk("begin"))
    para._p.append(_instr())
    para._p.append(_mk("separate"))
    # 佔位文字，Word 更新欄位後會被實際頁碼取代
    ph = para.add_run("–")
    ph.font.size = Pt(size)
    ph.bold = bold
    if color:
        ph.font.color.rgb = color
    ph.font.name = EN
    ph._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    para._p.append(_mk("end"))


def set_update_fields_on_open(doc):
    """讓 Word 開啟檔案時自動更新所有欄位（目錄頁碼、總頁數）"""
    settings = doc.settings.element
    for tag in ("w:updateFields",):
        existing = settings.find(qn(tag))
        if existing is not None:
            settings.remove(existing)
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    settings.append(el)


# ════════════════════════════════════════════════════════════════
#  版面元件
# ════════════════════════════════════════════════════════════════
def banner(doc, text, sub=None, bm=None, new_page=False):
    """章節大標,樣式由風格包的 banner 決定:
       filled  主色底反白字    rule  粗底線、無底色    sidebar  淺底加左側色條    plain  細底線
    字級用規格層的 h1。
    bm       = 書籤名稱,供目錄頁碼參照
    new_page = 於本段落前分頁。採 page_break_before 屬性而非插入分頁符段落,
               避免多出一個只裝分頁符的空段落而產生整頁空白。"""
    if new_page:
        strip_trailing_blanks(doc)
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if new_page:
        p.paragraph_format.page_break_before = True
    spc(p, before=180, after=120)
    h1, small = SZ["h1"], SZ["body"]
    if BANNER_STYLE == "rule":
        para_border(p, "bottom", 18, NAVY, space=4)
        run(p, text, bold=True, size=h1, color=NAVY, font=ZH_HEAD)
        if sub:
            run(p, f"   {sub}", bold=False, size=small, color=GRAY)
    elif BANNER_STYLE == "sidebar":
        para_bg(p, PALE)
        para_border(p, "left", 36, NAVY, space=8)
        run(p, f"  {text}", bold=True, size=h1, color=NAVY, font=ZH_HEAD)
        if sub:
            run(p, f"   {sub}", bold=False, size=small, color=GRAY)
    elif BANNER_STYLE == "plain":
        para_border(p, "bottom", 4, LINE, space=4)
        run(p, text, bold=True, size=h1, color=NAVY, font=ZH_HEAD)
        if sub:
            run(p, f"   {sub}", bold=False, size=small, color=GRAY)
    else:  # filled
        para_bg(p, NAVY)
        run(p, f"  {text}", bold=True, size=h1, color=WHITE, font=ZH_HEAD)
        if sub:
            run(p, f"   |   {sub}", bold=False, size=small, color=SUB_TINT)
    keep_with_next(p)
    if bm:
        add_bookmark(p, bm)
    return p


def subh(doc, text, num=None, new_page=False, bm=None):
    """次標題：鋼藍字 + 下框線。new_page 時於本段落前分頁"""
    if new_page:
        strip_trailing_blanks(doc)
    p = doc.add_paragraph(style="Heading 2")
    if new_page:
        p.paragraph_format.page_break_before = True
    spc(p, before=170, after=70)
    prefix = f"{num}  " if num else ("" if num == "" else "◆  ")
    r = p.add_run(prefix + text)
    r.bold = True
    r.font.size = Pt(SZ["h2"])
    r.font.color.rgb = STEEL
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH_HEAD)
    para_border(p, "bottom", 4, STEEL, space=2)
    keep_with_next(p)
    if bm:
        add_bookmark(p, bm)
    return p


def subsubh(doc, text, marker="▸  "):
    """三級標題:無框線,青色,預設前綴 ▸(招標文件規定標號格式時傳 marker="")"""
    p = doc.add_paragraph(style="Heading 3")
    spc(p, before=120, after=40)
    r = p.add_run(marker + text)
    r.bold = True
    r.font.size = Pt(SZ["h3"])
    r.font.color.rgb = TEAL
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH_HEAD)
    keep_with_next(p)
    return p


def body(doc, text, indent=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None):
    p = doc.add_paragraph()
    p.alignment = align
    spc(p, before=60, after=60, lr=LINE_RULE, ln=LINE_TW)
    r = p.add_run(text)
    r.font.size = Pt(size if size is not None else SZ["body"])
    r.font.color.rgb = INK
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    if indent is None:
        indent = FIRST_INDENT > 0
    if indent:
        p.paragraph_format.first_line_indent = Pt(SZ["body"] * (FIRST_INDENT or 2))
    return p


def bullet(doc, text, level=0):
    style = "List Bullet" if level == 0 else "List Bullet 2"
    p = doc.add_paragraph(style=style)
    spc(p, before=40, after=40, lr=LINE_RULE, ln=LINE_TW if LINE_RULE == "exact" else max(240, LINE_TW - 40))
    r = p.add_run(text)
    r.font.size = Pt(SZ["body"])
    r.font.color.rgb = INK
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    return p


def numlist(doc, text, n):
    """編號清單。編號以文字繪製並搭配凸排，
    不使用 Word「List Number」樣式——該樣式在同一份文件內會跨段落累加編號，
    合併多份文件後會出現 11.12.13... 之類的錯誤序號。"""
    p = doc.add_paragraph()
    spc(p, before=40, after=40, lr=LINE_RULE, ln=(LINE_TW if LINE_RULE == "exact" else 320))
    pf = p.paragraph_format
    pf.left_indent = Cm(0.95)
    pf.first_line_indent = Cm(-0.95)
    r = p.add_run(f"{n}.")
    r.bold = True
    r.font.size = Pt(SZ["body"])
    r.font.color.rgb = NAVY
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    r2 = p.add_run("　" + text)
    r2.font.size = Pt(SZ["body"])
    r2.font.color.rgb = INK
    r2.font.name = EN
    r2._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    return p


def numlist_block(doc, items, start=1):
    """一次輸出整組編號清單，序號自 start 起算"""
    for i, t in enumerate(items, start):
        numlist(doc, t, i)


def hbox(doc, text, bg=None, bar=None, color=None):
    """重點色塊：左粗邊 + 淺底"""
    bg = bg or PALE
    bar = bar or NAVY
    color = color or NAVY
    p = doc.add_paragraph()
    spc(p, before=100, after=100, lr=LINE_RULE, ln=(LINE_TW if LINE_RULE == "exact" else 340))
    para_bg(p, bg)
    para_border(p, "left", 24, bar, space=6)
    r = p.add_run(f"  {text}")
    r.font.size = Pt(SZ["body"])
    r.font.color.rgb = color
    r.bold = False
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    return p


def quotebox(doc, title, text):
    """標題 + 內文的強調框（用於承諾、原則宣告）"""
    p = doc.add_paragraph()
    spc(p, before=110, after=0)
    para_bg(p, PALE)
    para_border(p, "left", 24, TEAL, space=6)
    r = p.add_run(f"  {title}")
    r.bold = True
    r.font.size = Pt(SZ["body"])
    r.font.color.rgb = NAVY
    r.font.name = EN
    r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    spc(p2, before=0, after=110, lr=LINE_RULE, ln=(LINE_TW if LINE_RULE == "exact" else 330))
    para_bg(p2, PALE)
    para_border(p2, "left", 24, TEAL, space=6)
    r2 = p2.add_run(f"  {text}")
    r2.font.size = Pt(SZ["body"])
    r2.font.color.rgb = INK
    r2.font.name = EN
    r2._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)


def caption(doc, text, label=None):
    """圖說或表說。label 例如 "圖" 或 "表" 時,自動加上 SEQ 欄位編號(Word 可做交叉引用)。"""
    p = doc.add_paragraph(style="Caption")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spc(p, before=20, after=110)
    if label:
        _SEQ[label] = _SEQ.get(label, 0) + 1
        run(p, f"{label} ", size=SZ["caption"], color=GRAY)
        _complex_field(p, f"SEQ {label} \\* ARABIC", cached=str(_SEQ[label]), size=SZ["caption"], color=GRAY)
        run(p, f"　{text}", size=SZ["caption"], color=GRAY)
    else:
        run(p, text, size=SZ["caption"], color=GRAY)
    return p


def spacer(doc, pts=6):
    p = doc.add_paragraph()
    spc(p, before=0, after=pts * 20)
    return p


# ════════════════════════════════════════════════════════════════
#  表格
# ════════════════════════════════════════════════════════════════
_SUBTOTAL_KEYS = ["小計", "合計", "稅金", "總計", "首年合計", "續年合計"]
_CAT_PREFIX = ["一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、",
               "九、", "十、", "壹、", "貳、", "參、", "肆、", "伍、", "陸、"]


def tbl(doc, headers, rows, cw=None, size=None, center_from=None,
        head_bg=None, zebra=True):
    """
    通用表格。
      center_from : 從第 n 欄起置中（None = 全部靠左，除表頭）
      自動辨識分類列（首欄以「一、」開頭）與小計列並套色
    """
    size = size if size is not None else SZ["table"]
    if TABLE_STYLE == "zebra":
        head_bg, head_fg = head_bg or NAVY, WHITE
    elif TABLE_STYLE == "lines":
        head_bg, head_fg = head_bg or PALE, NAVY
        zebra = False
    else:  # minimal
        head_bg, head_fg = head_bg or WHITE, NAVY
        zebra = False
    n = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表頭
    for i, h in enumerate(t.rows[0].cells):
        set_cell_bg(h, head_bg)
        h.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spc(p, before=40, after=40)
        r = p.add_run(str(headers[i]))
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = head_fg
        r.font.name = EN
        r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)

    # 內容
    for ri, rd in enumerate(rows):
        rowcells = t.rows[ri + 1].cells
        first = str(rd[0]) if len(rd) else ""
        joined = "".join(str(x) for x in rd)
        is_cat = any(first.startswith(c) for c in _CAT_PREFIX)
        is_sub = any(k in joined for k in _SUBTOTAL_KEYS)
        base_bg = MIST if (zebra and ri % 2 == 0) else WHITE
        cell_bg = PALE if is_cat else (PALE if is_sub else base_bg)

        for ci in range(n):
            cell = rowcells[ci]
            set_cell_bg(cell, cell_bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            if center_from is not None and ci >= center_from:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif ci == 0 and n > 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            spc(p, before=40, after=40, lr=LINE_RULE, ln=(LINE_TW if LINE_RULE == "exact" else 300))

            txt = str(rd[ci]) if ci < len(rd) else ""
            # 支援儲存格內換行
            lines = txt.split("\n")
            for li, line in enumerate(lines):
                if li > 0:
                    p.add_run().add_break()
                r = p.add_run(line)
                r.bold = is_cat or is_sub
                r.font.size = Pt(size)
                r.font.color.rgb = NAVY if (is_cat or is_sub) else INK
                r.font.name = EN
                r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)

    if cw:
        t.autofit = False
        for row in t.rows:
            for ci, w in enumerate(cw):
                if ci < n:
                    row.cells[ci].width = Cm(w)
    style_table(t)
    doc.add_paragraph()
    return t


def kv_table(doc, pairs, cw=(4.0, 12.0), size=None):
    size = size if size is not None else SZ["table"]
    """兩欄式資料表：左欄淺底標籤、右欄內容"""
    t = doc.add_table(rows=len(pairs), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(pairs):
        c0, c1 = t.rows[i].cells
        set_cell_bg(c0, PALE)
        set_cell_bg(c1, WHITE)
        for cell, txt, bold, col, align in (
            (c0, k, True, NAVY, WD_ALIGN_PARAGRAPH.LEFT),
            (c1, v, False, INK, WD_ALIGN_PARAGRAPH.LEFT),
        ):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = align
            spc(p, before=40, after=40, lr=LINE_RULE, ln=(LINE_TW if LINE_RULE == "exact" else 300))
            for li, line in enumerate(str(txt).split("\n")):
                if li > 0:
                    p.add_run().add_break()
                r = p.add_run(line)
                r.bold = bold
                r.font.size = Pt(size)
                r.font.color.rgb = col
                r.font.name = EN
                r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    t.autofit = False
    for row in t.rows:
        row.cells[0].width = Cm(cw[0])
        row.cells[1].width = Cm(cw[1])
    set_table_borders(t, top=(4, str(LINE)), bottom=(8, str(NAVY)), insideH=(4, str(LINE)))
    table_rows_flow(t, header_rows=0)
    doc.add_paragraph()
    return t


def add_signature(para, width_cm=4.2, prefix="　　"):
    """在段落中插入負責人簽名檔（去背 PNG）"""
    if not os.path.exists(SIGNATURE):
        run(para, "　＿＿＿＿＿＿＿＿＿＿＿＿", bold=False, color=GRAY)
        return False
    if prefix:
        run(para, prefix, bold=False, color=INK)
    para.add_run().add_picture(SIGNATURE, width=Cm(width_cm))
    return True


def add_seals(para, company_cm=2.2, owner_cm=1.9, prefix="　　　"):
    """插入公司大章與負責人小章（依慣例大章在前、小章在後）"""
    ok = False
    if prefix:
        run(para, prefix, bold=False, color=INK)
    if os.path.exists(SEAL_COMPANY):
        para.add_run().add_picture(SEAL_COMPANY, width=Cm(company_cm))
        ok = True
    if os.path.exists(SEAL_OWNER):
        run(para, "　", bold=False, color=INK)
        para.add_run().add_picture(SEAL_OWNER, width=Cm(owner_cm))
        ok = True
    return ok


def sign_block(doc, label_date=True, seals=False):
    """投標廠商簽章區（附件一、附件二、服務建議書附錄共用）
    全區以 keepNext 綁定，避免被分頁拆散或落單成孤兒頁。"""
    spacer(doc, 6)
    for lbl, val in [
        ("投標廠商名稱：", BIDDER["name"]),
        ("統一編號：", BIDDER["tax_id"]),
        ("聯絡電話：", BIDDER["phone"]),
        ("電子郵件：", BIDDER["email"]),
    ]:
        p = doc.add_paragraph()
        spc(p, before=40, after=40)
        run(p, lbl, bold=True, color=NAVY)
        run(p, "  " + val, bold=False, color=INK)
        keep_with_next(p)

    # 負責人簽章：簽名 + 公司大章 + 負責人小章
    p = doc.add_paragraph()
    spc(p, before=70, after=40)
    run(p, "負責人簽章：", bold=True, color=NAVY)
    if seals:
        # 簽名與印章一起由 seals 控制;未經使用者確認時留白,不插入任何簽章圖檔
        add_signature(p, width_cm=4.0)
        add_seals(p)
    else:
        run(p, "　＿＿＿＿＿＿＿＿＿＿＿＿", bold=False, color=GRAY)
    if label_date:
        keep_with_next(p)

    if label_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spc(p, before=90, after=40)
        run(p, SIGN_DATE or "中　華　民　國　　　年　　　月　　　日",
            bold=False, color=INK)


# ════════════════════════════════════════════════════════════════
#  圖表
# ════════════════════════════════════════════════════════════════
def chart_img(doc, fig, w=15.5, cap=None, label=None):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spc(p, before=80, after=40)
    p.add_run().add_picture(buf, width=Cm(w))
    if cap:
        caption(doc, cap, label=label)


# ════════════════════════════════════════════════════════════════
#  頁首頁尾 / 頁碼
# ════════════════════════════════════════════════════════════════
def _field(para, instr):
    r = para.add_run()
    fld = OxmlElement("w:fldChar"); fld.set(qn("w:fldCharType"), "begin")
    r._r.append(fld)
    r2 = para.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = instr
    r2._r.append(it)
    r3 = para.add_run()
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r3._r.append(f2)
    for rr in (r, r2, r3):
        rr.font.size = Pt(9)
        rr.font.color.rgb = GRAY
        rr.font.name = EN
        rr._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)


def add_header_footer(doc, header_text=None, footer_left=None, skip_first=False):
    """頁首:左側案名簡稱、右側目前章名(STYLEREF,Word 開檔或輸出 PDF 時更新);頁尾:公司名(盲評不顯示)與頁碼。"""
    header_text = header_text or CASE_SHORT
    footer_left = footer_left if footer_left is not None else ("" if BLIND else BIDDER["name"])
    from docx.enum.text import WD_TAB_ALIGNMENT
    for section in doc.sections:
        if skip_first:  # 封面頁不顯示頁首頁尾
            section.different_first_page_header_footer = True
        tw = Cm(text_width_cm())
        if SHOW_HEADER and header_text:
            hp = section.header.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            hp.clear() if hasattr(hp, "clear") else None
            hp.paragraph_format.tab_stops.add_tab_stop(tw, WD_TAB_ALIGNMENT.RIGHT)
            r = hp.add_run(header_text)
            r.font.size = Pt(9)
            r.font.color.rgb = GRAY
            r.font.name = EN
            r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
            if HEADER_CHAPTER:
                hp.add_run("\t")
                _complex_field(hp, 'STYLEREF 1', cached="", size=9, color=GRAY)
            para_border(hp, "bottom", 4, LINE, space=2)
        if SHOW_PAGENUM:
            fp = section.footer.paragraphs[0]
            fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            fp.paragraph_format.tab_stops.add_tab_stop(tw, WD_TAB_ALIGNMENT.RIGHT)
            para_border(fp, "top", 4, LINE, space=2)
            if footer_left:
                r = fp.add_run(footer_left)
                r.font.size = Pt(9)
                r.font.color.rgb = GRAY
                r.font.name = EN
                r._element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
            fp.add_run("\t")
            _complex_field(fp, "PAGE", cached="1", size=9, color=NAVY, bold=True)
            r2 = fp.add_run(" / ")
            r2.font.size = Pt(9); r2.font.color.rgb = GRAY; r2.font.name = EN
            _complex_field(fp, "NUMPAGES", cached="1", size=9, color=GRAY)


def setup_page(doc, top=None, bottom=None, left=None, right=None):
    """依規格層設定紙張、方向、邊界與裝訂邊。"""
    w, h = PAPERS.get(PAPER, PAPERS["A4"])
    if LANDSCAPE:
        w, h = h, w
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(w), Cm(h)
        if LANDSCAPE:
            sec.orientation = WD_ORIENT.LANDSCAPE
        sec.top_margin = Cm(top if top is not None else MARGINS["top"])
        sec.bottom_margin = Cm(bottom if bottom is not None else MARGINS["bottom"])
        sec.left_margin = Cm(left if left is not None else MARGINS["left"])
        sec.right_margin = Cm(right if right is not None else MARGINS["right"])
        sec.gutter = Cm(GUTTER_CM)


def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = EN
    st.font.size = Pt(SZ["body"])
    st.element.rPr.rFonts.set(qn("w:eastAsia"), ZH)
    setup_page(doc)
    _SEQ.clear()
    define_styles(doc)
    return doc


def _cover_visual(doc):
    """封面視覺:使用者提供的主視覺圖優先;否則依風格包的 art 用色票自動生成幾何圖(種子取自案名)。"""
    tw = text_width_cm()
    if COVER_IMAGE and os.path.exists(COVER_IMAGE):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if COVER_LAYOUT == "center" else WD_ALIGN_PARAGRAPH.LEFT
        spc(p, before=0, after=120)
        p.add_run().add_picture(COVER_IMAGE, width=Cm(min(14, tw)))
        return
    if COVER_ART and COVER_ART != "none":
        import bid_components as C
        png = C.cover_art(COVER_ART, CASE_TITLE or "case", tw, 5.2)
        if png:
            p = doc.add_paragraph()
            spc(p, before=0, after=160)
            p.add_run().add_picture(io.BytesIO(png), width=Cm(tw))


def _cover_info_rows():
    """封面的廠商資訊;盲評案不顯示。"""
    if BLIND:
        return []
    rows = [("投標廠商", BIDDER["name"], True), ("統一編號", BIDDER["tax_id"], False),
            ("負責人", BIDDER["owner"], False)]
    if PARTNER and PARTNER.get("name"):
        rows.append(("協力單位", PARTNER["name"], False))
    return rows


def build_cover(doc, doc_kind="服務建議書", doc_title=None, doc_sub=None):
    """封面。版式由風格包的 cover 決定:center 置中、left 靠左、band 頂部色帶。
    盲評案(configure(blind=True))不顯示公司名、統編、負責人、協力單位。"""
    doc_title = doc_title or doc_kind
    title = CASE_TITLE or doc_title
    ct = SZ["cover_title"]
    sm = SZ["body"]
    spacer(doc, 26 if COVER_LAYOUT != "band" else 8)

    if COVER_LAYOUT != "band":
        _cover_visual(doc)

    al = WD_ALIGN_PARAGRAPH.CENTER if COVER_LAYOUT == "center" else WD_ALIGN_PARAGRAPH.LEFT

    if COVER_LAYOUT == "band":
        t = doc.add_table(rows=1, cols=1)
        c = t.rows[0].cells[0]
        set_cell_bg(c, NAVY)
        c.paragraphs[0].text = ""
        pp = c.paragraphs[0]
        spc(pp, before=240, after=60)
        if CLIENT:
            run(pp, CLIENT, size=SZ["body"] + 1, color=SUB_TINT, font=ZH_HEAD)
        p2 = c.add_paragraph()
        spc(p2, before=120, after=60)
        run(p2, title, bold=True, size=ct, color=WHITE, font=ZH_HEAD)
        p3 = c.add_paragraph()
        spc(p3, before=0, after=240)
        run(p3, doc_kind + (f"   {CASE_SUBTITLE}" if CASE_SUBTITLE else ""), bold=False,
            size=SZ["body"] + 4, color=SUB_TINT, font=ZH_HEAD)
        spacer(doc, 20)
    else:
        if CLIENT:
            p = doc.add_paragraph(); p.alignment = al
            spc(p, before=0, after=60)
            run(p, CLIENT, bold=False, size=SZ["body"] + 2, color=STEEL, font=ZH_HEAD)
        if CASE_SUBTITLE and COVER_LAYOUT == "center":
            p = doc.add_paragraph(); p.alignment = al
            spc(p, before=0, after=200)
            run(p, CASE_SUBTITLE, bold=False, size=SZ["body"] + 2, color=STEEL, font=ZH_HEAD)
        p = doc.add_paragraph(); p.alignment = al
        spc(p, before=60 if COVER_LAYOUT == "left" else 0, after=60)
        run(p, title, bold=True, size=ct, color=NAVY, font=ZH_HEAD)
        p = doc.add_paragraph(); p.alignment = al
        spc(p, before=0, after=40)
        run(p, doc_kind, bold=True, size=SZ["body"] + 7, color=ACCENT if COVER_LAYOUT == "left" else STEEL, font=ZH_HEAD)
        if CASE_SUBTITLE and COVER_LAYOUT == "left":
            p = doc.add_paragraph(); p.alignment = al
            spc(p, before=0, after=40)
            run(p, CASE_SUBTITLE, bold=False, size=SZ["body"], color=GRAY)
        if doc_sub:
            p = doc.add_paragraph(); p.alignment = al
            spc(p, before=0, after=40)
            run(p, doc_sub, bold=False, size=SZ["body"] + 1, color=GRAY)
        p = doc.add_paragraph(); p.alignment = al
        spc(p, before=120, after=260)
        para_border(p, "bottom", 12, NAVY, space=1)
        spacer(doc, 14)

    for label, value, strong in _cover_info_rows():
        p = doc.add_paragraph(); p.alignment = al
        spc(p, before=30, after=30)
        run(p, f"{label}　　", bold=False, size=sm, color=GRAY)
        run(p, value, bold=strong, size=SZ["body"] + (4 if strong else 0), color=NAVY if strong else INK)

    spacer(doc, 16)
    if COVER_DATE:
        p = doc.add_paragraph(); p.alignment = al
        spc(p, before=0, after=0)
        run(p, COVER_DATE, bold=False, size=SZ["body"], color=STEEL)
    if COVER_LAYOUT == "band":
        spacer(doc, 30)
        _cover_visual(doc)


if __name__ == "__main__":
    print("bid_kit 為共用模組,請由 build_proposal.py 或自己的 gen_*.py 匯入。")
