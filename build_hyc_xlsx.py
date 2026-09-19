#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_hyc_xlsx.py — 从李家村底表 XML 级克隆出华阳城底表（保留全部公式/条件格式/样式）

思路与 write_xs_xml.py 一致：直接改 xlsx 内部 XML，绝不用 openpyxl 整本重写
（openpyxl 会丢条件格式样式）。

改动：
  1) 全 XML 字符串替换：门店名 / 店员名 / 店长名
  2) 月度任务 sheet：A/B 列人员 + C..N 任务值（晨哥 2026-09-19 给的《华阳城合作店员工销售数据》）
     另加 O 列「滞销(参考)」

用法: python3 build_hyc_xlsx.py
"""
import os, re, shutil, subprocess, sys, zipfile
import xml.etree.ElementTree as ET

SRC = "/Users/mac/Desktop/李家村销售/李家村9月任务进度.xlsx"
OUTDIR = "/Users/mac/Desktop/华阳城销售"
OUT = os.path.join(OUTDIR, "华阳城9月任务进度.xlsx")
WORK = "/tmp/hyc_build"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def Q(name):
    return "{%s}%s" % (NS, name)
ET.register_namespace("", NS)

# ---------- 替换表（顺序敏感：长串在前） ----------
REPL = [
    ("零售发展部-华为李家村万达授权体验店", "华为华阳城合作店"),
    ("华为李家村万达授权体验店", "华为华阳城合作店"),
    ("华为事业部龙湖店店长绩效考评表", "华为事业部华阳城合作店店长绩效考评表"),
    ("李家村", "华阳城"),
    ("邵乐乐", "张梅B"),
    ("杨丽华", "李俊琪"),
    ("李泽", "王莹莹B"),
    ("陈超磊", "张赫桐"),
    ("张博晨", "田蕊"),
]

# ---------- 华阳城月度任务（晨哥 2026-09-19 提供） ----------
# 列: A店面 B销售 C收入(销额) D毛利 E手机 F增值服务 G合约 HPC I平板 J音频 K穿戴 LHD(智慧屏) M摄影课程 N考核机
PEOPLE = ["张梅B", "李俊琪", "王莹莹B", "张赫桐"]
TASKS = {   # 姓名 -> {列字母: 值}
    "张梅B":   {"C": 575000, "D": 82500, "E": 55, "F": 20750, "G": 4, "H": 5, "I": 9,  "J": 11, "K": 33, "L": 1, "M": 0, "N": 0, "O": 4},
    "李俊琪":  {"C": 575000, "D": 82500, "E": 55, "F": 20750, "G": 4, "H": 5, "I": 9,  "J": 12, "K": 32, "L": 1, "M": 0, "N": 0, "O": 4},
    "王莹莹B": {"C": 575000, "D": 82500, "E": 55, "F": 20750, "G": 4, "H": 5, "I": 10, "J": 11, "K": 33, "L": 1, "M": 0, "N": 0, "O": 4},
    "张赫桐":  {"C": 575000, "D": 82500, "E": 55, "F": 20750, "G": 4, "H": 5, "I": 10, "J": 11, "K": 32, "L": 0, "M": 0, "N": 0, "O": 4},
}
ROWS = {4: "张梅B", 5: "李俊琪", 6: "王莹莹B", 7: "张赫桐"}


def col_of(ref):
    return "".join(ch for ch in ref if ch.isalpha())


def sheet_files(root):
    """返回 [(sheet文件绝对路径, sheet名)]，顺序按 workbook.xml。"""
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    wb = ET.parse(os.path.join(root, "xl/workbook.xml")).getroot()
    relroot = ET.parse(os.path.join(root, "xl/_rels/workbook.xml.rels")).getroot()
    rid2t = {}
    for rel in relroot:
        rid2t[rel.get("Id")] = rel.get("Target")
    sheets_el = None
    for el in wb:
        if el.tag.endswith("}sheets"):
            sheets_el = el
            break
    out = []
    for sh in (sheets_el if sheets_el is not None else []):
        name = sh.get("name")
        rid = sh.get(f"{R}id")
        t = (rid2t.get(rid) or "").lstrip("/")
        if not t:
            continue
        if not t.startswith("xl/"):
            t = "xl/" + t
        out.append((os.path.join(root, t), name))
    return out


def str_replace_all(root):
    n = 0
    for dirpath, _, files in os.walk(root):
        for fn in files:
            if not fn.endswith((".xml", ".rels")):
                continue
            p = os.path.join(dirpath, fn)
            s = open(p, encoding="utf-8").read()
            o = s
            for a, b in REPL:
                s = s.replace(a, b)
            if s != o:
                open(p, "w", encoding="utf-8").write(s)
                n += 1
    return n


def _child(el, localname):
    for c in el:
        if c.tag.endswith("}" + localname) or c.tag == localname:
            return c
    return None


def set_cell(ws_root, ref, value, is_str=False):
    """在 worksheet XML 中设置单元格值（新建或覆盖）。"""
    row_idx = int(re.sub(r"\D", "", ref))
    sd = _child(ws_root, "sheetData")
    if sd is None:
        sys.exit("❌ sheet 无 sheetData")
    tr = None
    for r in sd:
        if r.get("r") == str(row_idx):
            tr = r
            break
    if tr is None:
        sys.exit(f"❌ 找不到行 {row_idx}")
    c = None
    for cc in tr:
        if cc.get("r") == ref:
            c = cc
            break
    if c is None:
        c = ET.SubElement(tr, Q("c"))
        c.set("r", ref)
        kids = sorted(list(tr), key=lambda e: (len(col_of(e.get("r", "A"))), col_of(e.get("r", "A"))))
        for ch in list(tr):
            tr.remove(ch)
        for ch in kids:
            tr.append(ch)
    for ch in list(c):
        c.remove(ch)
    if is_str:
        c.set("t", "inlineStr")
        is_el = ET.SubElement(c, Q("is"))
        t = ET.SubElement(is_el, Q("t"))
        t.text = value
    else:
        if "t" in c.attrib:
            del c.attrib["t"]
        v = ET.SubElement(c, Q("v"))
        v.text = str(value)


def main():
    if not os.path.exists(SRC):
        sys.exit(f"❌ 源底表不存在: {SRC}")
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    with zipfile.ZipFile(SRC) as z:
        z.extractall(WORK)

    changed = str_replace_all(WORK)
    print(f"✓ 字符串替换完成（{changed} 个 XML 文件）")

    # 月度任务 = 第 1 个 sheet
    sheets = sheet_files(WORK)
    tk_path, tk_name = sheets[0]
    print(f"  月度任务 sheet 文件: {os.path.relpath(tk_path, WORK)}  名={tk_name}")

    tree = ET.parse(tk_path)
    root = tree.getroot()
    # 表头 O3
    set_cell(root, "O3", "滞销(参考)", is_str=True)
    for r, name in ROWS.items():
        set_cell(root, f"B{r}", name, is_str=True)
        for col, val in TASKS[name].items():
            set_cell(root, f"{col}{r}", val)
    set_cell(root, "O8", "合计", is_str=True)
    # A8 合计行店面名
    set_cell(root, "A8", "华为华阳城合作店", is_str=True)
    tree.write(tk_path, encoding="utf-8", xml_declaration=True)
    print("✓ 月度任务：4 名店员 + 任务值 + 滞销(参考) 列已写入")

    byname = {nm: p for p, nm in sheets}

    # 渠道挂账：第 7 行补「张赫桐」（原陈超磊行已空）+ 任务额
    qd_path = byname.get("渠道挂账")
    t2 = ET.parse(qd_path); r2 = t2.getroot()
    set_cell(r2, "A7", "张赫桐", is_str=True)
    set_cell(r2, "B7", 41400)
    t2.write(qd_path, encoding="utf-8", xml_declaration=True)
    print("✓ 渠道挂账：A7=张赫桐 B7=41400")

    # 华阳城销售：清空乐机收手工项 T14:V18（李家村旧值不带过来），T19/U19 SUM 公式保留
    sl_path = byname.get("华阳城销售")
    t3 = ET.parse(sl_path); r3 = t3.getroot()
    for row in range(14, 19):
        for col in ("T", "U", "V"):
            try:
                set_cell(r3, f"{col}{row}", 0)
            except SystemExit:
                pass
    t3.write(sl_path, encoding="utf-8", xml_declaration=True)
    print("✓ 华阳城销售：乐机收 T14:V18 已清零")

    # 打包
    os.makedirs(OUTDIR, exist_ok=True)
    if os.path.exists(OUT):
        os.remove(OUT)
    # [Content_Types].xml 必须第一个
    allfiles = []
    for dirpath, _, files in os.walk(WORK):
        for fn in files:
            p = os.path.join(dirpath, fn)
            allfiles.append(os.path.relpath(p, WORK))
    allfiles.sort(key=lambda x: (x != "[Content_Types].xml",))
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for rel in allfiles:
            z.write(os.path.join(WORK, rel), rel)
    print(f"✓ 已生成: {OUT}  ({os.path.getsize(OUT)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
