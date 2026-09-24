#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""relogin_mgr_py.py — 用友【经理号】无头重登（playwright 版，移植自李家村 mem17 relogin_yonyou.py）。
与 bash 版 relogin_mgr.sh 差异：元素用确定性 ID 选择器（#username/#password/#submit_btn_login）、
等待更长（30s+5s 兜底两轮找 iframe）、不依赖 accessibility snapshot 文案，偶发"表单未出现"概率大幅降低。
输出 ~/.agent-browser/sessions/yonyou-mgr-default.json（{"cookies":[...]}，与 agent-browser state 同构，
fetch_sales_analysis_hyc.load_cookies 直接可用）。密码钥匙串 service=yonyou-mgr，绝不落日志。"""
import time, json, subprocess, os, sys
from playwright.sync_api import sync_playwright

CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
YY = "https://c3.yonyoucloud.com"
ACCT = "18591910491"
STATE = os.path.expanduser("~/.agent-browser/sessions/yonyou-mgr-default.json")

pwd = subprocess.run(["security", "find-generic-password", "-s", "yonyou-mgr", "-w"],
                     capture_output=True, text=True, timeout=15).stdout.strip()
if not pwd:
    raise SystemExit("ERR: 钥匙串未找到 yonyou-mgr 密码")

with sync_playwright() as p:
    b = p.chromium.launch(headless=True, executable_path=CHROME,
                          args=["--no-sandbox", "--disable-dev-shm-usage",
                                "--disable-blink-features=AutomationControlled"])
    ctx = b.new_context()
    pg = ctx.new_page()
    pg.goto(YY, wait_until="domcontentloaded", timeout=30000)
    time.sleep(3)
    try:
        btn = pg.query_selector(".button_accept")
        if btn: btn.click(); time.sleep(1)
    except Exception: pass
    lf = None
    for _ in range(2):                      # 两轮找登录 iframe（含 5s 兜底）
        for f in pg.frames:
            if "euc.yonyoucloud.com" in (f.url or ""): lf = f; break
        if lf: break
        time.sleep(5)
    if not lf:
        b.close(); raise SystemExit("NO_LOGIN_FRAME")
    lf.fill("#username", ACCT); time.sleep(0.3)
    lf.fill("#password", pwd); time.sleep(0.3)
    lf.click("#submit_btn_login")
    ok = False
    for _ in range(15):
        time.sleep(2)
        if "login" not in pg.url.lower() and "cas" not in pg.url.lower():
            ok = True; break
    time.sleep(3)
    if not ok:
        b.close(); raise SystemExit("LOGIN_TIMEOUT url=" + pg.url)
    cookies = ctx.cookies()
    names = {c.get("name") for c in cookies}
    if "yht_access_token" not in names:
        b.close(); raise SystemExit("NO_TOKEN: 登录后缺少 yht_access_token")
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    if os.path.exists(STATE):
        try:
            import shutil; shutil.copyfile(STATE, STATE + ".bak")
        except Exception: pass
    json.dump({"cookies": cookies}, open(STATE, "w"), ensure_ascii=False)
    n = len([c for c in cookies if "yonyou" in c.get("domain", "")])
    b.close()
    print("LOGGED_IN cookies=%d (yonyou=%d) -> %s" % (len(cookies), n, STATE))
