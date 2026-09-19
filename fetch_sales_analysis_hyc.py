#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_sales_analysis_hyc.py —— 拉「华为华阳城合作店」的销售分析明细（经理账号）。

为什么单独写：
  1) 李家村用店长账号(18161914293)，返回天然只含李家村，无需门店过滤；
     华阳城必须用**经理账号**(18591910491)，该账号返回**全公司全部门店**，
     必须在本地按 `store_name == "华为华阳城合作店"` 过滤。
     ⚠️ 历史血泪：曾误判"经理号只看得到大唐不夜城"——真相是 report/list
     结果按门店排序，第 1 页 5000 行恰好全是第一家店。别只看第一页就下结论。
  2) rm_saleanalysis 的日期/门店/org 等 queryParams **服务端全部无效**，
     必须全量分页拉取（约 46 页 × 5000）后本地过滤。
  3) 全公司数据量大（22.8 万行），逐页过滤而非全量驻留内存，降低占用。

产出：
  sa_hyc_month.json   —— 本月（当月1日~今天）华阳城记录（全 124 字段，已去重）
  sa_aug_cache.json   —— 同上的 merge_qudao.py 消费格式（{records:[...]}），
                         供渠道明细/员工×渠道/渠道达成使用
  sa_warehouse_hyc.json —— 华阳城全历史（当月切片的上游，供复用/审计）
  sa_raw.tsv          —— 本月宽表（人工核对用）

用法：
  python fetch_sales_analysis_hyc.py [--use-cache] [--month 2026-09]
"""
import json, os, sys, ssl, time, datetime, urllib.request, urllib.error, re
import collections
from concurrent.futures import ThreadPoolExecutor, as_completed

YY_BASE = "https://c3.yonyoucloud.com"
SA_URL = YY_BASE + "/yonbip-mkt-retailweb/report/list"
BILLNUM = "rm_saleanalysis"

STORE_NAME = "华为华阳城合作店"
STORE_KEY = "华阳城"          # 宽松匹配键（防门店名写法微调）
SN_FIELD = "oid_userDefine_2419863036093267976"  # 序列号

BASE = os.path.dirname(os.path.abspath(__file__))
# 经理账号登录态（店长账号看不到华阳城）
DEFAULT_STATE = os.path.expanduser("~/.agent-browser/sessions/yonyou-mgr-default.json")
STATE = os.environ.get("YONYOU_STATE", DEFAULT_STATE)

OUT_MONTH = os.path.join(BASE, "sa_hyc_month.json")
OUT_CACHE = os.path.join(BASE, "sa_aug_cache.json")
OUT_WAREHOUSE = os.path.join(BASE, "sa_warehouse_hyc.json")
OUT_TSV = os.path.join(BASE, "sa_raw.tsv")

PAGE_SIZE = 5000
MAX_WORKERS = 5          # 并发页（服务端单页约 60s，串行 46 页要 45 分钟）
CACHE_MAX_AGE = 6 * 3600

USE_CACHE = "--use-cache" in sys.argv
MONTH = None
if "--month" in sys.argv:
    MONTH = sys.argv[sys.argv.index("--month") + 1]

TODAY = datetime.date.today()
if MONTH:
    y, m = MONTH.split("-")
    MONTH_FIRST = datetime.date(int(y), int(m), 1)
else:
    MONTH_FIRST = TODAY.replace(day=1)
    MONTH = MONTH_FIRST.strftime("%Y-%m")


def load_cookies(path):
    d = json.load(open(path))
    return {c["name"]: c["value"] for c in d.get("cookies", [])
            if "yonyou" in c.get("domain", "") and c.get("name") and c.get("value") is not None}


def _clean(v):
    """HTTP 头只能是 latin-1 安全字符，cookie 里的非 ASCII 要剥掉。"""
    return re.sub(r"[^\x20-\x7e]", "", str(v))


def build_headers(ck):
    return {
        "User-Agent": "Mozilla/5.0 (Macintosh) AppleWebKit/537.36 Chrome/120 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": YY_BASE,
        "Referer": YY_BASE + "/",
        "Cookie": "; ".join("%s=%s" % (k, _clean(v)) for k, v in ck.items()),
        "XSRF-TOKEN": _clean(ck.get("XSRF-TOKEN", "")),
        "yht_access_token": _clean(ck.get("yht_access_token", "")),
    }


def fetch_page(hdr, page_index, retries=3, timeout=240):
    body = json.dumps({
        "billnum": BILLNUM,
        "page": {"pageIndex": page_index, "pageSize": PAGE_SIZE},
    }).encode("utf-8")
    last = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(SA_URL, data=body, headers=hdr, method="POST")
            with urllib.request.urlopen(req, timeout=timeout,
                                        context=ssl.create_default_context()) as resp:
                j = json.loads(resp.read())
            if j.get("code") != 200:
                raise RuntimeError("接口 code=%s msg=%s" % (j.get("code"), str(j.get("message"))[:120]))
            dd = j.get("data") or {}
            return dd.get("recordCount"), (dd.get("recordList") or [])
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise SystemExit("✗ HTTP_%d 登录态失效，请先跑 relogin_mgr.sh 重登经理账号" % e.code)
            last = e
        except SystemExit:
            raise
        except Exception as e:
            last = e
        if attempt < retries:
            time.sleep(2 * attempt)
    raise RuntimeError("页 %d 拉取失败: %s" % (page_index, last))


def is_hyc(rec):
    s = str(rec.get("store_name") or "")
    return STORE_KEY in s


def dedup(records):
    """剔除 _YD 预订单 + 按 (单号, SKU, 序列号) 去重（与李家村管线同口径）。"""
    seen, out = set(), []
    for r in records:
        code = str(r.get("code") or "")
        if code.endswith("_YD"):
            continue
        k = (code, str(r.get("productsku_cCode") or ""), str(r.get(SN_FIELD) or ""))
        if k in seen:
            continue
        seen.add(k)
        out.append(r)
    return out


def atomic_write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)
    os.replace(tmp, path)


def main():
    if not os.path.exists(STATE):
        sys.exit("✗ 找不到经理账号登录态: %s（先跑 relogin_mgr.sh）" % STATE)
    ck = load_cookies(STATE)
    if "yht_access_token" not in ck:
        sys.exit("✗ 登录态缺少 yht_access_token，请重登")
    hdr = build_headers(ck)

    print("▶ 拉取销售分析（经理号 %s）→ 本地过滤 store_name 含 '%s'"
          % (os.path.basename(STATE), STORE_KEY))
    print("  目标月份: %s（%s ~ %s）" % (MONTH, MONTH_FIRST, TODAY))

    # ---- 仓模式 ----
    hyc_all = []
    if USE_CACHE and os.path.exists(OUT_WAREHOUSE):
        try:
            wh = json.load(open(OUT_WAREHOUSE, encoding="utf-8"))
            age = time.time() - wh.get("fetched_at", 0)
            if age < CACHE_MAX_AGE:
                hyc_all = wh["records"]
                print("  [仓] 读本地仓（%.0f 分钟前，%d 行），跳过联网" % (age / 60, len(hyc_all)))
            else:
                print("  [仓] 已过期(%.1f h)，转联网" % (age / 3600))
        except Exception as e:
            print("  [仓] 读取失败，转联网: %s" % e)

    if not hyc_all:
        t0 = time.time()
        total, page1 = fetch_page(hdr, 1)
        n_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        print("  接口 recordCount=%s（全公司）→ %d 页 × %d" % (total, n_pages, PAGE_SIZE))
        hyc_all.extend(r for r in page1 if is_hyc(r))
        print("  [页 1/%d] 本页华阳城 %d 行，累计 %d" % (n_pages, len(hyc_all), len(hyc_all)))

        done = 1
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futs = {pool.submit(fetch_page, hdr, pi): pi for pi in range(2, n_pages + 1)}
            for fut in as_completed(futs):
                pi = futs[fut]
                try:
                    _, rl = fut.result()
                    got = [r for r in rl if is_hyc(r)]
                    hyc_all.extend(got)
                    done += 1
                    print("  [页 %d/%d] 本页华阳城 %d 行，累计 %d（%.0fs）"
                          % (pi, n_pages, len(got), len(hyc_all), time.time() - t0))
                except Exception as e:
                    print("  [页 %d] 失败: %s" % (pi, e))

        raw_n = len(hyc_all)
        hyc_all = dedup(hyc_all)
        print("  [去重] %d → %d 行" % (raw_n, len(hyc_all)))
        try:
            atomic_write_json(OUT_WAREHOUSE, {"fetched_at": time.time(),
                                              "store": STORE_NAME, "records": hyc_all})
            print("  [仓] 已写全历史仓 %d 行" % len(hyc_all))
        except Exception as e:
            print("  [仓] 写仓失败(不影响本次): %s" % e)

    # ---- 本地按业务日期筛本月 ----
    begin = MONTH_FIRST.strftime("%Y-%m-%d")
    end = TODAY.strftime("%Y-%m-%d")
    month_rows = [r for r in hyc_all if begin <= str(r.get("dDate") or "")[:10] <= end]
    print("  [筛选] 按 dDate %s~%s 保留 %d / %d 行" % (begin, end, len(month_rows), len(hyc_all)))

    if not month_rows:
        sys.exit("✗ 本月无华阳城记录，检查门店名/日期口径")

    # 规范化 dDate 为 ISO（与 update_sa_cache 口径一致）
    for r in month_rows:
        d = str(r.get("dDate") or "")[:10]
        if d:
            r["dDate"] = d

    atomic_write_json(OUT_MONTH, {"begin": begin, "end": end,
                                 "recordCount": len(month_rows), "records": month_rows})
    atomic_write_json(OUT_CACHE, {"records": month_rows})

    # 宽表 TSV（人工核对）
    cols, seen = [], set()
    for r in month_rows:
        for k in r:
            if k not in seen:
                seen.add(k)
                cols.append(k)
    with open(OUT_TSV, "w", encoding="utf-8") as f:
        f.write("\t".join(cols) + "\n")
        for r in month_rows:
            f.write("\t".join(str(r.get(c, "")) for c in cols) + "\n")

    net = sum(float(r.get("fNetMoney") or 0) for r in month_rows)
    qty = sum(float(r.get("fQuantity") or 0) for r in month_rows)
    emps = collections.Counter(r.get("iEmployeeid_name") for r in month_rows)
    print("✓ 华阳城本月销售分析已落盘")
    print("  记录 %d 行 | 销售净额 ¥%.2f | 数量 %.0f" % (len(month_rows), net, qty))
    print("  业务员分布: %s" % dict(emps))
    print("  → %s" % OUT_MONTH)
    print("  → %s（merge_qudao 消费）" % OUT_CACHE)
    return 0


if __name__ == "__main__":
    sys.exit(main())
