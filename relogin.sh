#!/usr/bin/env bash
# relogin.sh —— 用友重新登录 + 导出登录态（供 fetch_yonyou_http.py 纯HTTP使用）
# 背景：fetch_yonyou_http.py 读 ~/.agent-browser/sessions/yonyou-default.json 里的
#       yht_access_token。该 token 过期后看板流水线报 HTTP_500/401。
#       本脚本在浏览器内完成登录，并把登录态导出覆盖到 yonyou-default.json。
# 用法: ./relogin.sh
# 退出码: 0 成功 / 1 失败
set -uo pipefail

export AGENT_BROWSER_EXECUTABLE_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
export AGENT_BROWSER_SESSION_NAME="yonyou"
AB="/Users/mac/.workbuddy/binaries/node/workspace/node_modules/.bin/agent-browser"
PY="/Users/mac/.workbuddy/binaries/python/versions/3.13.12/bin/python3"
SESSION_FILE="$HOME/.agent-browser/sessions/yonyou-default.json"
ACCOUNT="18161914293"
BASE="https://c3.yonyoucloud.com"
unset HTTP_PROXY HTTPS_PROXY ALL_PROXY

PASS="$(security find-generic-password -s yonyou -w 2>/dev/null)"
[ -z "$PASS" ] && { echo "✗ 钥匙串没有用友密码"; exit 1; }

log(){ echo -e "\n\033[1;36m[$(date '+%H:%M:%S')] $*\033[0m"; }

log "打开用友…"
"$AB" close --all >/dev/null 2>&1 || true
"$AB" open "$BASE/#/" >/dev/null 2>&1

# 等待登录表单稳定（iframe yonbip_login 加载）
ACC_REF=""
for i in 1 2 3 4 5 6 7 8 9 10; do
  sleep 2
  SNAP="$("$AB" snapshot 2>/dev/null)"
  ACC_REF="$(echo "$SNAP" | grep '邮箱/账号/用户手机号' | grep -o 'ref=e[0-9]*' | head -1 | cut -d= -f2)"
  [ -n "$ACC_REF" ] && break
done
if [ -z "$ACC_REF" ]; then
  echo "✗ 等待超时：登录表单未出现（可能页面改版或出验证码）"
  exit 1
fi
PWD_REF="$(echo "$SNAP" | grep 'textbox "密码"' | grep -o 'ref=e[0-9]*' | head -1 | cut -d= -f2)"
BTN_REF="$(echo "$SNAP" | grep 'button "登录"' | grep -o 'ref=e[0-9]*' | head -1 | cut -d= -f2)"
log "表单就绪 acc=$ACC_REF pwd=$PWD_REF btn=$BTN_REF"

# 点掉 Cookie 弹窗（若存在）
"$AB" click "@e4" >/dev/null 2>&1 || true
sleep 1

"$AB" fill "@$ACC_REF" "$ACCOUNT" >/dev/null 2>&1
sleep 1
"$AB" fill "@$PWD_REF" "$PASS" >/dev/null 2>&1
sleep 1
"$AB" click "@$BTN_REF" >/dev/null 2>&1

# 等待登录结果（title 离开「数字化工作台」= 登录成功）
OK=""
for i in 1 2 3 4 5 6 7 8 9 10; do
  sleep 3
  TITLE="$("$AB" get title 2>/dev/null | tail -1)"
  if [ -n "$TITLE" ] && [ "$TITLE" != "数字化工作台" ]; then
    OK=1
    break
  fi
done
if [ -z "$OK" ]; then
  echo "✗ 登录未成功（title 仍为「数字化工作台」，可能验证码/账号异常）"
  exit 1
fi
log "登录成功: $TITLE"

# 导出登录态到会话文件（保持同一浏览器连接，立即导出）
cp -f "$SESSION_FILE" "$SESSION_FILE.bak" 2>/dev/null || true
"$AB" state save "$SESSION_FILE" >/dev/null 2>&1
"$PY" -c "
import json, sys
d = json.load(open('$SESSION_FILE'))
names = [c.get('name') for c in d.get('cookies', [])]
if 'yht_access_token' not in names:
    print('✗ 导出后仍缺少 yht_access_token，登录态未捕获'); sys.exit(1)
print('✓ 登录态已刷新: %d cookies（含 yht_access_token）' % len(names))
"
