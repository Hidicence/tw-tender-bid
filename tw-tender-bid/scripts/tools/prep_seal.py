# -*- coding: utf-8 -*-
"""
處理簽名與印章圖檔:白底轉透明、裁掉多餘留白,輸出可插入 Word 的 PNG。

用法:
    python -I prep_seal.py <來源圖> <輸出.png> [--ink 26,32,44] [--force] [--white-cut 205] [--ink-cut 120]

  --ink    把筆畫統一成這個 RGB 墨色(簽名常用深灰黑);印章通常不指定,保留原色
  --force  即使來源已有透明度也重做去背(掃描檔的筆畫周圍常殘留白色暈邊)

安全提醒:簽名與印章圖檔是敏感資料,不要放進公開 repo、雲端共用資料夾或 skill 資料夾;
只把路徑記在公司 profile 的 seals_and_signature。產出的成品也只在使用者確認後才插入文件。
"""
import argparse, os, sys
from PIL import Image


def process(src, out, ink_rgb=None, force=False, white_cut=205, ink_cut=120):
    im = Image.open(src).convert("RGBA")
    w0, h0 = im.size
    px = im.load()
    # 真的已去背:需有足夠比例的全透明像素(原始掃描檔常帶少量雜訊 alpha,不能只看最小 alpha)
    hist = im.getchannel("A").histogram()
    already = sum(hist[:16]) > (w0 * h0) * 0.10 and not force
    if not already:
        for y in range(im.size[1]):
            for x in range(im.size[0]):
                r, g, b, a = px[x, y]
                lum = (r * 299 + g * 587 + b * 114) // 1000
                if lum >= white_cut:
                    px[x, y] = (0, 0, 0, 0)
                elif lum <= ink_cut:
                    px[x, y] = (ink_rgb + (255,)) if ink_rgb else (r, g, b, 255)
                else:
                    alpha = max(0, min(255, int(255 * (white_cut - lum) / (white_cut - ink_cut))))
                    px[x, y] = (ink_rgb + (alpha,)) if ink_rgb else (r, g, b, alpha)
    bbox = im.getbbox()
    if bbox:
        im = im.crop(bbox)
    pad = max(3, im.size[0] // 70)
    canvas = Image.new("RGBA", (im.size[0] + pad * 2, im.size[1] + pad * 2), (0, 0, 0, 0))
    canvas.paste(im, (pad, pad), im)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    canvas.save(out)
    return w0, h0, canvas.size, already


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--ink", help="R,G,B")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--white-cut", type=int, default=205)
    ap.add_argument("--ink-cut", type=int, default=120)
    a = ap.parse_args()
    if not os.path.exists(a.src):
        sys.exit(f"找不到來源圖:{a.src}")
    ink = tuple(int(v) for v in a.ink.split(",")) if a.ink else None
    w0, h0, size, already = process(a.src, a.out, ink, a.force, a.white_cut, a.ink_cut)
    print(f"OK {w0}x{h0} -> {size[0]}x{size[1]}  比例 {size[0] / size[1]:.2f}  {'(原檔已去背)' if already else '(本次去背)'}")
    print("輸出:", a.out)


if __name__ == "__main__":
    main()
