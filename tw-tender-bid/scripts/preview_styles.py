# -*- coding: utf-8 -*-
"""
風格預覽:用同一份範例內容,依序套用多個風格包,各產生一份 Word(含封面、章標題、內文、表格、圖表),
並在有 Word 的 Windows 上輸出 PDF 與總覽圖,方便使用者挑風格。

用法:
    python -I preview_styles.py --out <資料夾> [--styles id1,id2,id3] [--format format.json] [--pdf]
    python -I preview_styles.py --list

--format 會套用招標文件的格式規格,預覽出來的就是「照規格後的樣子」。
--styles 不指定時,從 8 個風格包裡依「流派」各挑一個(最多 3 個),避免三個方向長得太像。
--avoid 傳入最近用過的風格 id,預設排除,讓每一案風格不同。
"""
import argparse, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bid_kit as K  # noqa: E402
from build_proposal import build  # noqa: E402

SAMPLE_SECTIONS = [
    {"type": "banner", "text": "壹、對需求的理解", "sub": "評審項目一"},
    {"type": "lead", "text": "我們把機關的需求拆成三件事:讓資訊找得到、讓流程看得懂、讓成果驗收得了。"},
    {"type": "stats", "items": [
        {"value": "12", "unit": "年", "label": "專案執行經驗"},
        {"value": "86", "unit": "件", "label": "類似案件實績"},
        {"value": "9", "unit": "人", "label": "核心團隊"},
        {"value": "100", "unit": "%", "label": "準時交付率"}]},
    {"type": "subh", "text": "背景與目的", "num": "1.1"},
    {"type": "body", "text": "本案旨在提升機關對外服務的清晰度與可信度。我們先釐清受眾、情境與限制,再把需求拆成可驗收的工作項目,每一項都對應一個交付物與審核點。"},
    {"type": "callout", "kind": "key", "title": "我們的承諾", "text": "每個里程碑都有機關確認點,未經確認不進入下一階段。"},
    {"type": "subh", "text": "工作方法", "num": "1.2"},
    {"type": "steps", "items": [
        {"title": "需求訪談與場勘", "text": "釐清受眾、使用情境與限制,產出需求確認書。"},
        {"title": "設計與製作", "text": "提出初稿,依機關意見修訂至多三輪。"},
        {"title": "驗收與交付", "text": "依驗收標準逐項檢核,交付成果與原始檔。"}]},
    {"type": "banner", "text": "貳、執行計畫", "sub": "評審項目二"},
    {"type": "twocol", "left": {"title": "現況", "items": ["資訊分散,民眾不易找到", "流程說明不一致"]},
     "right": {"title": "我們的做法", "items": ["整理成單一入口與清楚動線", "統一用語與流程圖"]}},
    {"type": "table", "caption": "工作項目與期程", "headers": ["階段", "工作內容", "交付物", "期程"],
     "rows": [["一、前期", "", "", ""], ["需求確認", "訪談與場勘", "需求確認書", "第 1 至 2 週"],
              ["製作", "設計與執行", "初稿與修訂稿", "第 3 至 8 週"], ["小計", "", "", "8 週"]],
     "widths": [3.0, 5.0, 4.5, 3.5]},
    {"type": "chart", "kind": "gantt", "unit": "週", "caption": "專案時程(示意)",
     "data": [{"task": "需求確認", "start": 0, "end": 2}, {"task": "設計", "start": 2, "end": 5},
              {"task": "執行", "start": 4, "end": 8}, {"task": "驗收", "start": 8, "end": 9}]},
]


def pick_styles(avoid):
    """依流派各挑一個,排除最近用過的。"""
    by_school = {}
    for sid, name, school, best in K.list_styles():
        if sid in avoid:
            continue
        by_school.setdefault(school, []).append(sid)
    picks = [v[0] for v in by_school.values()]
    return picks[:3]


def export_pdf(docx, pdf):
    try:
        import win32com.client
    except Exception:
        return False
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.Fields.Update()
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
    finally:
        word.Quit()
    return os.path.exists(pdf)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--out")
    ap.add_argument("--styles")
    ap.add_argument("--format")
    ap.add_argument("--avoid", default="")
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    if a.list:
        for sid, name, school, best in K.list_styles():
            print(f"{sid:20s} {name}  [{school}]  適合:{best}")
        return 0
    if not a.out:
        ap.error("--out 必填")
    os.makedirs(a.out, exist_ok=True)
    fmt = None
    if a.format:
        with open(a.format, encoding="utf-8") as f:
            fmt = json.load(f)
    avoid = set(x for x in a.avoid.split(",") if x)
    styles = [x for x in a.styles.split(",") if x] if a.styles else pick_styles(avoid)
    made = []
    for sid in styles:
        spec = {"doc_kind": "服務建議書", "style": sid, "format": fmt, "toc": True,
                "case": {"title": "範例機關 專案服務案", "short": "範例專案服務案", "client": "範例機關",
                         "subtitle": "案號 EX-000-000", "cover_date": "中華民國 115 年 11 月"},
                "sections": [dict(s) for s in SAMPLE_SECTIONS]}
        docx = os.path.join(a.out, f"preview_{sid}.docx")
        build(spec, docx)
        made.append(docx)
        if a.pdf:
            pdf = docx[:-5] + ".pdf"
            print(("PDF:" if export_pdf(docx, pdf) else "未能輸出 PDF(需要本機 Word):") + pdf)
    for m in made:
        print("預覽:", m)
    return 0


if __name__ == "__main__":
    sys.exit(main())
