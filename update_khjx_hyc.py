#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_khjx_hyc.py —— 就地填华阳城底表「月度任务」的「考核机」任务列（N 列）

为什么需要：
  底表是从李家村 XML 级克隆来的，而晨哥给的《华阳城合作店员工销售数据.xlsx》
  只有 13 列（…电脑/平板/智慧/滞销），**没有「考核机」这一列**，克隆时按"不臆造"
  填了 0。看板上于是显示「门店任务 0 台」、每人「任务 —」。
  2026-09-19 晨哥确认：**每人 4 台**（4 人共 16 台，店长田蕊不背）。

做法：
  XML 级原位改单元格（保留 s= 样式），不用 openpyxl 整本重写（会丢条件格式/样式）。
  只动「月度任务」sheet 的 N4:N7，其他一格不碰；公式单元格天然跳过。

用法：
  python update_khjx_hyc.py [底表.xlsx] [--value 4] [--dry]
"""
import os
import re
import sys
import datetime
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

DEFAULT_XLSX = "/Users/mac/Desktop/华阳城销售/华阳城9月任务进度.xlsx"
SHEET = "月度任务"
SALES_ROWS = [4, 5, 6, 7]          # 张梅B / 李俊琪 / 王莹莹B / 张赫桐（r8=合计，不动）
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def sheet_files(root):
    wb = ET.parse(os.path.join(root, "xl/workbook.xml")).getroot()
    relroot = ET.parse(os.path.join(root, "xl/_rels/workbook.xml.rels")).getroot()
    rid2t = {rel.get("Id"): rel.get("Target") for rel in relroot}
    sheets_el = None
    for el in wb:
        if el.tag.endswith("}sheets"):
            sheets_el = el
            break
    out = {}
    for sh in (sheets_el if sheets_el is not None else []):
        t = (rid2t.get(sh.get(f"{R_NS}id")) or "").lstrip("/")
        if not t:
            continue
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[sh.get("name")] = t
    return out


def set_num(xml, ref, val):
    """把单元格设为数值；保留 s= 样式；公式/不存在的单元格原样返回。"""
    m = re.search(r'<c r="%s"([^>]*?)(/>|>(.*?)</c>)' % re.escape(ref), xml, re.S)
    if not m:
        return xml, "absent"
    attrs = m.group(1) or ""
    inner = m.group(3) or ""
    if "<f" in inner or "<f" in attrs:
        return xml, "formula"
    s = re.search(r's="(\d+)"', attrs)
    new = '<c r="%s"%s><v>%s</v></c>' % (ref, (' s="%s"' % s.group(1)) if s else "", val)
    return xml[:m.start()] + new + xml[m.end():], "set"


def main():
    args = [a for a in sys.argv[1:]]
    dry = "--dry" in args
    val = 4
    if "--value" in args:
        val = args[args.index("--value") + 1]
    xlsx = next((a for a in args if a.endswith(".xlsx")), DEFAULT_XLSX)
    if not os.path.exists(xlsx):
        sys.exit(f"❌ 找不到底表: {xlsx}")

    work = tempfile.mkdtemp(prefix="hyc_khjx_")
    with zipfile.ZipFile(xlsx) as z:
        z.extractall(work)
    by = sheet_files(work)
    rel = by.get(SHEET)
    if not rel:
        sys.exit(f"❌ 底表里找不到 sheet「{SHEET}」")

    p = os.path.join(work, rel)
    xml = open(p, encoding="utf-8").read()
    before = {}
    for r in SALES_ROWS:
        ref = "N%d" % r
        m = re.search(r'<c r="%s"[^>]*>(.*?)</c>' % ref, xml, re.S)
        before[ref] = (re.search(r"<v>([^<]*)</v>", m.group(1)).group(1) if m and "<v>" in m.group(1) else "空")
    print("改前 N4:N7 =", before)

    cnt = {"set": 0, "formula": 0, "absent": 0}
    for r in SALES_ROWS:
        xml, st = set_num(xml, "N%d" % r, val)
        cnt[st] = cnt.get(st, 0) + 1
        print(f"  N{r} -> {val}  [{st}]")
    open(p, "w", encoding="utf-8").write(xml)

    if dry:
        print("(--dry 试跑，未写回)")
        shutil.rmtree(work, ignore_errors=True)
        return 0

    bak_dir = os.path.join(os.path.dirname(xlsx), "_备份")
    os.makedirs(bak_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(xlsx))[0]
    bak = os.path.join(bak_dir, f"{base}_考核机填值前{stamp}.xlsx")
    shutil.copy2(xlsx, bak)
    print(f"→ 已备份: {os.path.basename(bak)}")

    out = xlsx + ".tmp"
    files = []
    for dirpath, _, fns in os.walk(work):
        for fn in fns:
            files.append(os.path.relpath(os.path.join(dirpath, fn), work))
    files.sort(key=lambda x: (x != "[Content_Types].xml",))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(os.path.join(work, f), f)
    os.replace(out, xlsx)
    os.system('xattr -c "%s" 2>/dev/null' % xlsx)
    shutil.rmtree(work, ignore_errors=True)

    print(f"✓ 考核机任务已填入（每人 {val} 台）: {xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
