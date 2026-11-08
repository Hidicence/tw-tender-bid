# -*- coding: utf-8 -*-
"""
報價與金額工具:中文大寫金額、含稅計算、千分位格式。

用法(命令列):
    python -I price_utils.py 1062600                 印出千分位與中文大寫
    python -I price_utils.py 1012000 --tax 0.05      未稅轉含稅,並印出各項

用法(程式):
    import sys; sys.path.insert(0, "<skill>/scripts/tools")
    from price_utils import to_cn_amount, fmt, with_tax

注意:最終金額一律由使用者確認;這裡只做換算,不決定價格。
"""
import argparse, sys

_CN_NUM = "零壹貳參肆伍陸柒捌玖"
_CN_UNIT = ["", "拾", "佰", "仟"]
_CN_SEC = ["", "萬", "億"]


def to_cn_amount(n):
    """整數轉中文大寫(不含「元整」)。例:1062600 -> 壹佰零陸萬貳仟陸佰。"""
    if n is None:
        return "＿＿＿＿＿＿＿＿＿＿＿＿"
    n = int(n)
    if n == 0:
        return "零"
    secs, tmp = [], n
    while tmp > 0:
        secs.append(tmp % 10000)
        tmp //= 10000
    out, need_zero = "", False
    for si in range(len(secs) - 1, -1, -1):
        sec = secs[si]
        if sec == 0:
            need_zero = True
            continue
        # 前面已有數字時:中間有空段,或這一段不足四位(例如 10005、1000500),都要補一個「零」
        if out and (need_zero or sec < 1000):
            out += "零"
        need_zero = False
        s, unit, part = sec, 0, ""
        while s > 0:
            d = s % 10
            if d == 0:
                if part and not part.startswith("零"):
                    part = "零" + part
            else:
                part = _CN_NUM[d] + _CN_UNIT[unit] + part
            s //= 10
            unit += 1
        out += part + _CN_SEC[si]
    return out


def fmt(n):
    """千分位。"""
    return f"{int(n):,}"


def with_tax(net, rate=0.05):
    """未稅轉含稅(四捨五入到整數)。慣例上先算好每個項目的未稅整數,再乘 1 + 稅率。"""
    return int(round(net * (1 + rate)))


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("amount", type=int)
    ap.add_argument("--tax", type=float, help="稅率(例 0.05);給了就把輸入視為未稅金額並換算含稅")
    a = ap.parse_args()
    if a.tax is not None:
        gross = with_tax(a.amount, a.tax)
        print(f"未稅 {fmt(a.amount)}  稅額 {fmt(gross - a.amount)}  含稅 {fmt(gross)}")
        print(f"含稅大寫:新台幣{to_cn_amount(gross)}元整")
    else:
        print(f"{fmt(a.amount)}  新台幣{to_cn_amount(a.amount)}元整")


if __name__ == "__main__":
    main()
