#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reset_manual_hyc.py —— 清空华阳城底表里「从李家村克隆过来的手工打分项」

为什么必须做：
  build_hyc_xlsx.py 是 XML 级克隆，公式/条件格式全保真，但**手工填写的常量也会被带过来**：
    · 个人表 L12:L17（投诉 / 执行力 / MA大考 / 晨读 / 全科生加减分）—— 本是
      邵乐乐 / 杨丽华 / 李泽 / 陈超磊 的李家村月度手工分；
    · 店长表 J11:J23（运营指标 / 达人标签 / 企业文化 / 超库龄加减分）—— 本是
      店长张博晨的李家村月度手工分。
  若不清空，华阳城看板会把李家村某人得的分算到华阳城店员头上（实测影响：
  李俊琪 +10、张赫桐 +10、王莹莹B +5、张梅B +2；田蕊 30.50 全为张博晨的分）。
  → 原则：**不臆造**。手工项一律清零，等晨哥按华阳城实际情况填入。

只动这两类常量单元格；公式单元格（L4:L11 / J4:J10）一个不碰。
XML 级原位替换，不用 openpyxl 整本重写（会丢条件格式/样式）。

用法：
  python reset_manual_hyc.py [底表.xlsx]
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

SALES = ["张梅B", "李俊琪", "王莹莹B", "张赫桐"]
INDIVIDUAL_REFS = [f"L{r}" for r in range(12, 18)]          # 个人表手工项
MANAGER_REFS = [f"J{r}" for r in range(11, 24)]              # 店长表手工项（J4:J10 是公式，天然跳过）

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def sheet_files(root):
    """返回 {sheet名: 包内相对路径}。"""
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


def blank_cell(xml, ref):
    """把某单元格置空（保留 s= 样式，去掉 t= 与内容）；公式单元格原样返回。"""
    m = re.search(r'<c r="%s"([^>]*?)(/>|>.*?</c>)' % re.escape(ref), xml, re.S)
    if not m:
        return xml, "absent"
    attrs = m.group(1) or ""
    body = m.group(2) or ""
    if "<f" in body or "<f" in attrs:
        return xml, "formula"
    s = re.search(r's="(\d+)"', attrs)
    new = '<c r="%s"%s/>' % (ref, (' s="%s"' % s.group(1)) if s else "")
    return xml[:m.start()] + new + xml[m.end():], "blanked"


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    if not os.path.exists(xlsx):
        sys.exit(f"❌ 找不到底表: {xlsx}")

    work = tempfile.mkdtemp(prefix="hyc_reset_")
    with zipfile.ZipFile(xlsx) as z:
        z.extractall(work)
    by = sheet_files(work)

    targets = [(n, INDIVIDUAL_REFS) for n in SALES] + [("店长", MANAGER_REFS)]
    stat = {}
    for sname, refs in targets:
        rel = by.get(sname)
        if not rel:
            print(f"  ⚠️ 找不到 sheet: {sname}")
            continue
        p = os.path.join(work, rel)
        xml = open(p, encoding="utf-8").read()
        cnt = {"blanked": 0, "formula": 0, "absent": 0}
        for ref in refs:
            xml, st = blank_cell(xml, ref)
            cnt[st] = cnt.get(st, 0) + 1
        open(p, "w", encoding="utf-8").write(xml)
        stat[sname] = cnt
        print(f"  {sname:<8} 已清空 {cnt['blanked']} 格 | 保留公式 {cnt['formula']} 格 | 原本为空 {cnt['absent']} 格")

    # 备份 + 重新打包
    bak_dir = os.path.join(os.path.dirname(xlsx), "_备份")
    os.makedirs(bak_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(xlsx))[0]
    bak = os.path.join(bak_dir, f"{base}_手工项清零前{stamp}.xlsx")
    shutil.copy2(xlsx, bak)
    print(f"→ 已备份: {os.path.basename(bak)}")

    out = xlsx + ".tmp"
    files = []
    for dirpath, _, fns in os.walk(work):
        for fn in fns:
            p = os.path.join(dirpath, fn)
            files.append(os.path.relpath(p, work))
    files.sort(key=lambda x: (x != "[Content_Types].xml",))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for rel in files:
            z.write(os.path.join(work, rel), rel)
    os.replace(out, xlsx)
    os.system('xattr -c "%s" 2>/dev/null' % xlsx)
    shutil.rmtree(work, ignore_errors=True)

    print(f"✓ 手工项已清零: {xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
