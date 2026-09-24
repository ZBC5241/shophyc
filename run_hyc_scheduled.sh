#!/bin/bash
# ============================================================
# run_hyc_scheduled.sh —— 华阳城看板「定时更新」入口（由 launchd 调用）
#
# 职责：设好环境变量（playwright 的 NODE_PATH 必须有，否则拉数脚本找不到模块）
#       调用 update_hyc.sh 全流程，并统一收口日志。
#
# 排期由 com.hyc.dashboard.update.plist 控制（9 档/天，2026-09-20 调整）：
#   14:12 / 18:12 / 20:12 / 20:37 / 21:07 / 21:37 / 21:50 / 22:07 / 22:55
#
# 【2026-09-21 自愈三件套移植（17号，晨哥拍板）】：
#   ① 失败重试：失败后 60 秒自动重试 1 次（覆盖接口偶发/重登偶发）
#   ② 企微告警：仍失败则 notify_fail.py 发看板群（不再静默断更）
#   ③ 登录态保鲜：成功后把经理号登录态同步到 data/state/ 稳定副本
#
# 并发保护：update_hyc.sh 内置 mkdir 互斥锁，档位重叠时后者自动跳过
#           （21:50 档约跑 3 分钟，与 22:07 档间隔 17 分钟，无冲突）。
#
# 手动跑：bash run_hyc_scheduled.sh
# ============================================================
set -u

HERE="$(cd "$(dirname "$0")" && pwd)"
export NODE_PATH="/Users/mac/.workbuddy/binaries/node/workspace/node_modules"
export PATH="/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

LOG_DIR="$HERE/logs"
LOG="$LOG_DIR/hyc_scheduled.log"
mkdir -p "$LOG_DIR"

# 日志轮转：超过 5MB 只留最后 3000 行，避免长期堆积（晨哥 2026-09-20：减轻储存压力）
if [ -f "$LOG" ] && [ "$(stat -f%z "$LOG" 2>/dev/null || echo 0)" -gt 5242880 ]; then
  tail -3000 "$LOG" > "$LOG.tmp" && mv -f "$LOG.tmp" "$LOG"
  echo "[log] 已轮转（保留最后 3000 行）" >> "$LOG"
fi

{
  echo ""
  echo "════════════════════════════════════════════════════════════"
  echo "  华阳城看板 · 定时更新  $(date '+%Y-%m-%d %H:%M:%S')"
  echo "════════════════════════════════════════════════════════════"
} >> "$LOG"

bash "$HERE/update_hyc.sh" >> "$LOG" 2>&1
RC=$?

# ① 失败重试（60 秒后 1 次；互斥锁天然防与下一档重叠）
if [ "$RC" -ne 0 ]; then
  echo "[retry] 首次失败 RC=$RC，60 秒后自动重试 1 次" >> "$LOG"
  sleep 60
  bash "$HERE/update_hyc.sh" >> "$LOG" 2>&1
  RC=$?
  echo "[$(date '+%H:%M:%S')] 重试退出码 $RC" >> "$LOG"
fi

# ② 仍失败 → 企微告警（发「🤖李家村数据看板」群，复用李家村 notify_fail.py）
if [ "$RC" -ne 0 ]; then
  /usr/bin/python3 /Users/mac/.local/share/TeleAgent/TeleAgent的工作空间/shop/notify_fail.py \
    --stage "华阳城看板定时更新(shophyc·17号自愈壳)" --exit-code "$RC" --log "$LOG" --tail 15 >> "$LOG" 2>&1 \
    || echo "[alert] 告警发送失败（不影响退出码）" >> "$LOG"
else
  # ③ 登录态保鲜：成功后把经理号登录态同步到稳定副本（仿 mem17）
  MGR_STATE="$HOME/.agent-browser/sessions/yonyou-mgr-default.json"
  STABLE="$HERE/data/state/yonyou-mgr_state.json"
  mkdir -p "$HERE/data/state"
  if [ -f "$MGR_STATE" ]; then cp -f "$MGR_STATE" "$STABLE" && echo "[keepalive] 登录态已保鲜 -> data/state/" >> "$LOG"; fi
fi

echo "[$(date '+%H:%M:%S')] 退出码 $RC" >> "$LOG"
exit $RC
