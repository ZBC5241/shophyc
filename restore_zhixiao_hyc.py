#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
restore_zhixiao_hyc.py —— 修复「WPS 另存底表」造成的两处损伤

背景（2026-09-19 20:44 查明）：
  华阳城底表在 17:13 之后被 **WPS 打开并保存过**（证据：包内出现
  xmlns:dbsheet="http://web.wps.cn/et/2021/dbsheet"、
  xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"）。
  WPS 另存带来两个副作用：
    ① 「月度任务」sheet 的 **O 列「滞销(参考)」整列被清空**
       （O3 表头 / O4:O6 的值 / O8「合计」全丢，只有 O7 侥幸留下）；
    ② workbook.xml 的 `<calcPr fullCalcOnLoad="1"/>` 被简化成 `<calcPr calcId="191029"/>`，
       底表失去「打开即全量重算」标记，且部分公式缓存值会停在旧值
       （如 N8 合计 = 12，实际 SUM(N4:N7) = 16）。

本脚本只做三件事，其余一格不碰：
  1. 恢复 O3 = 滞销(参考)、O4:O7 = 4、O8 = 合计（值取自 17:13 建表版备份）
  2. 修正 N8（考核机合计）的**缓存值** = SUM(N4:N7)，使其与 N 列一致
  3. 补回 calcPr 的 fullCalcOnLoad="1"

文本一律写成 inlineStr，避免动 sharedStrings.xml。
XML 级原位改，不用 openpyxl 整本重写（会再洗一遍样式）。

用法：
  python restore_zhixiao_hyc.py [底表.xlsx] [--dry]
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
ZHIXIAO_VALUE = 4                     # 晨哥给的任务表「滞销」列 = 4/4/4/4
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


def cell_re(ref):
    return re.compile(r'<c r="%s"((?:\s[^>]*?)?)(?:/>|>(.*?)</c>)' % re.escape(ref), re.S)


def read_cell(xml, ref):
    m = cell_re(ref).search(xml)
    if not m:
        return None, "absent"
    inner = m.group(2)
    if inner is None:
        return None, "empty"
    if "<f" in inner:
        v = re.search(r"<v>([^<]*)</v>", inner)
        return (v.group(1) if v else None), "formula"
    v = re.search(r"<v>([^<]*)</v>", inner)
    if v:
        return v.group(1), "value"
    t = re.search(r"<t[^>]*>([^<]*)</t>", inner)
    return (t.group(1) if t else None), "value"


def style_of(xml, ref):
    m = cell_re(ref).search(xml)
    if not m:
        return None
    s = re.search(r'\bs="(\d+)"', m.group(1) or "")
    return s.group(1) if s else None


def put_num(xml, ref, val):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, "absent"
    if m.group(2) is not None and "<f" in m.group(2):
        return xml, "formula"
    s = style_of(xml, ref)
    new = '<c r="%s"%s><v>%s</v></c>' % (ref, (' s="%s"' % s) if s else "", val)
    return xml[:m.start()] + new + xml[m.end():], "set"


def put_str(xml, ref, text):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, "absent"
    if m.group(2) is not None and "<f" in m.group(2):
        return xml, "formula"
    s = style_of(xml, ref)
    esc = (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    new = '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, (' s="%s"' % s) if s else "", esc)
    return xml[:m.start()] + new + xml[m.end():], "set"


def patch_calcpr(wbxml):
    """补回 fullCalcOnLoad="1"。"""
    m = re.search(r"<calcPr\b[^>]*/>", wbxml)
    if not m:
        return wbxml, "absent"
    tag = m.group(0)
    if "fullCalcOnLoad" in tag:
        return wbxml, "already"
    new = tag[:-2].rstrip() + ' fullCalcOnLoad="1"/>'
    return wbxml[:m.start()] + new + wbxml[m.end():], "set"


def main():
    args = sys.argv[1:]
    dry = "--dry" in args
    xlsx = next((a for a in args if a.endswith(".xlsx")), DEFAULT_XLSX)
    if not os.path.exists(xlsx):
        sys.exit(f"❌ 找不到底表: {xlsx}")

    work = tempfile.mkdtemp(prefix="hyc_fix_")
    with zipfile.ZipFile(xlsx) as z:
        z.extractall(work)
    by = sheet_files(work)
    rel = by.get(SHEET)
    if not rel:
        sys.exit(f"❌ 找不到 sheet「{SHEET}」")

    # ---- 1. 滞销列 ----
    p = os.path.join(work, rel)
    xml = open(p, encoding="utf-8").read()
    print("[1] 恢复「滞销(参考)」列")
    for ref, want in [("O3", "滞销(参考)"), ("O4", ZHIXIAO_VALUE), ("O5", ZHIXIAO_VALUE),
                      ("O6", ZHIXIAO_VALUE), ("O7", ZHIXIAO_VALUE), ("O8", "合计")]:
        old, st = read_cell(xml, ref)
        if isinstance(want, int):
            xml, r = put_num(xml, ref, want)
        else:
            xml, r = put_str(xml, ref, want)
        print(f"    {ref}: {old!r}({st}) -> {want!r}  [{r}]")

    # ---- 2. 考核机合计缓存值 ----
    vals = []
    for r in range(4, 8):
        v, _ = read_cell(xml, "N%d" % r)
        try:
            vals.append(float(v))
        except (TypeError, ValueError):
            vals.append(0.0)
    total = sum(vals)
    old_n8, st_n8 = read_cell(xml, "N8")
    # N8 是公式格：只改 <v> 缓存值
    m = cell_re("N8").search(xml)
    if m and m.group(2) and "<f" in m.group(2):
        inner = re.sub(r"<v>[^<]*</v>", "", m.group(2))
        inner = inner.replace("<v/>", "")
        inner_new = inner + "<v>%g</v>" % total
        xml = xml[:m.start()] + '<c r="N8"%s>%s</c>' % (m.group(1) or "", inner_new) + xml[m.end():]
        print(f"[2] N8 合计缓存值: {old_n8} -> {total:g}  (N4:N7 = {vals})")
    else:
        print(f"[2] ⚠️ N8 不是公式格，跳过（读到的值 {old_n8!r}/{st_n8}）")
    open(p, "w", encoding="utf-8").write(xml)

    # ---- 3. fullCalcOnLoad ----
    wbp = os.path.join(work, "xl/workbook.xml")
    wbxml = open(wbp, encoding="utf-8").read()
    before = re.search(r"<calcPr\b[^>]*/>", wbxml)
    wbxml, r = patch_calcpr(wbxml)
    open(wbp, "w", encoding="utf-8").write(wbxml)
    after = re.search(r"<calcPr\b[^>]*/>", wbxml)
    print("[3] calcPr: %s -> %s  [%s]" % (before.group(0) if before else "无",
                                          after.group(0) if after else "无", r))

    if dry:
        print("(--dry 试跑，未写回)")
        shutil.rmtree(work, ignore_errors=True)
        return 0

    bak_dir = os.path.join(os.path.dirname(xlsx), "_备份")
    os.makedirs(bak_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(xlsx))[0]
    bak = os.path.join(bak_dir, f"{base}_WPS损伤修复前{stamp}.xlsx")
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
    print(f"✓ 修复完成: {xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
