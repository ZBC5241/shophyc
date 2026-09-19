#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_hyc.py —— 华阳城看板数字对照校验（独立复算底表 SUMIFS）

不依赖 Excel：直接按《华阳城销售》sheet 的真实公式语义，从过滤后的
「华阳城门店毛利明细表」独立复算每人每项，再与 data.json 逐格对照。

覆盖《华阳城销售》!C/G/K/O/S/W/AA/AM 列口径（行 4~8，即 4 名销售 + 田蕊）：
  毛利 C = SUMIFS(XS!N, XS!P=姓名)
  手机 G = SUMIFS(XS!I, XS!P=姓名, XS!D="01手机")
  PC   K = ... D="05电脑"
  平板 O = ... D="06平板电脑"
  穿戴 S = ... D="08穿戴"
  音频 W = ... D="07音频"
  HD  AA = SUMIFS(XS!I, XS!P=姓名, XS!F="12*")      # SKU 编码前缀 12
  销额 AM = SUMIFS(XS!M, XS!P=姓名)

用法：
  python verify_hyc.py <华阳城门店毛利明细.xlsx> [data.json]
"""
import sys, os, json, collections
import warnings
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = sys.argv[1] if len(sys.argv) > 1 else \
    "/Users/mac/.local/share/TeleAgent/playwright-mcp/hyc/华阳城门店毛利明细表-华为终端.xlsx"
DATA = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "data.json")

PEOPLE = ["张梅B", "李俊琪", "王莹莹B", "张赫桐", "田蕊"]
METRICS = [
    ("毛利", None, None, 13),      # Σ 毛利列(14)
    ("手机", 3, "01手机", 8),
    ("PC", 3, "05电脑", 8),
    ("平板", 3, "06平板电脑", 8),
    ("穿戴", 3, "08穿戴", 8),
    ("音频", 3, "07音频", 8),
    ("销额", None, None, 12),      # Σ 金额列(13)
]


def main():
    if not os.path.exists(XLSX):
        sys.exit(f"❌ 找不到明细: {XLSX}")
    warnings.filterwarnings("ignore", message="Workbook contains no default style")
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    next(it)
    seen = set()
    agg = {n: collections.Counter() for n in PEOPLE}
    n_rows = 0
    for r in it:
        if not r or not r[0] or not str(r[0]).strip():
            continue
        key = (str(r[0]).strip(), str(r[5]).strip())   # 出库单号 + SKU 编码
        if key in seen:
            continue
        seen.add(key)
        n_rows += 1
        who = str(r[15] or "").strip()
        if who not in agg:
            continue
        cat = str(r[3] or "").strip()
        sku = str(r[5] or "").strip()

        def f(i):
            try:
                return float(str(r[i]).replace(",", "") or 0)
            except (TypeError, ValueError):
                return 0.0

        agg[who]["毛利"] += f(13)
        agg[who]["销额"] += f(12)
        if cat == "01手机":
            agg[who]["手机"] += f(8)
        if cat == "05电脑":
            agg[who]["PC"] += f(8)
        if cat == "06平板电脑":
            agg[who]["平板"] += f(8)
        if cat == "08穿戴":
            agg[who]["穿戴"] += f(8)
        if cat == "07音频":
            agg[who]["音频"] += f(8)
        if sku.startswith("12"):
            agg[who]["HD"] += f(8)

    data = json.load(open(DATA, encoding="utf-8"))
    store_perf = data["store"]["performance"]
    ppl = data["people"]

    print(f"明细去重 {n_rows} 行  ←→  data.json  ({data['meta']['date']})\n")
    bad = 0
    for m, catidx, catval, col in METRICS:
        print(f"── {m} ──")
        tot_a = tot_b = 0.0
        for n in PEOPLE:
            a = round(agg[n][m], 2)
            seg = (ppl.get(n, {}).get("performance", {}) or {}).get(m) or {}
            b = seg.get("done")
            b = None if b is None else round(float(b), 2)
            tot_a += a
            tot_b += (b or 0)
            ok = (b is not None and abs(a - b) < 0.02)
            if not ok:
                bad += 1
            print(f"   {n:<8} 独立复算 {a:>12,.2f}   data.json {('%.2f' % b) if b is not None else '—':>12}   {'OK' if ok else '❌ 不一致'}")
        # 门店合计再与 store.performance 对照
        st_done = (store_perf.get(m) or {}).get("done")
        st_done = None if st_done is None else round(float(st_done), 2)
        ok_tot = st_done is not None and abs(tot_a - st_done) < 0.02
        if not ok_tot:
            bad += 1
        print(f"   {'合计':<8} 独立复算 {tot_a:>12,.2f}   store表 {('%.2f' % st_done) if st_done is not None else '—':>12}   {'OK' if ok_tot else '❌'}\n")

    print("── 门店合计 ──")
    gm = round(sum(agg[n]['毛利'] for n in PEOPLE), 2)
    print(f"   毛利合计 独立复算 {gm:,.2f}")

    print("\n" + ("❌ 存在 %d 项不一致" % bad if bad else "✅ 全部指标与底表公式口径逐格一致"))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
