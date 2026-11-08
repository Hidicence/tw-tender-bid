# -*- coding: utf-8 -*-
"""
把投標文件的各份 PDF 依清單順序合併成單一 PDF,並建立書籤(目次)與頁數摘要。

用法:
    python -I merge_pdf.py <manifest.json>

manifest 格式見 assets/submission.manifest.example.json。合併順序請照招標文件「投標文件清單」的順序。
電子投標常要求「全彩 PDF」「單一檔案」「郵件主旨格式」,這些以須知為準,這裡只負責合併。
"""
import json, os, sys
from pypdf import PdfReader, PdfWriter


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    mpath = os.path.abspath(sys.argv[1])
    base = os.path.dirname(mpath)
    with open(mpath, encoding="utf-8") as f:
        m = json.load(f)

    def resolve(p):
        return p if os.path.isabs(p) else os.path.join(m.get("base_dir") or base, p)

    out = resolve(m["output_pdf"])
    w, cursor, rows = PdfWriter(), 0, []
    for part in m["parts"]:
        path = resolve(part["pdf"])
        if not os.path.exists(path):
            sys.exit(f"[缺檔] {path}")
        r = PdfReader(path)
        for pg in r.pages:
            w.add_page(pg)
        w.add_outline_item(part["title"], cursor)
        rows.append((part["title"], cursor + 1, cursor + len(r.pages), len(r.pages)))
        cursor += len(r.pages)
    meta = m.get("metadata") or {}
    if meta:
        w.add_metadata({("/" + k): v for k, v in meta.items()})
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "wb") as f:
        w.write(f)
    print("OK ->", out)
    print(f"{'文件':<40}{'起':>6}{'迄':>6}{'頁數':>6}")
    print("-" * 58)
    for t, s, e, n in rows:
        print(f"{t:<40}{s:>6}{e:>6}{n:>6}")
    print("-" * 58)
    print(f"{'合計':<40}{'':>6}{'':>6}{cursor:>6}\n檔案大小:{os.path.getsize(out) / 1024 / 1024:.2f} MB")


if __name__ == "__main__":
    main()
