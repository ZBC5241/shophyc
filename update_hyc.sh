#!/bin/bash
# ============================================================
# update_hyc.sh —— 华阳城看板「更新看板」一键全流程
#
# 与李家村 shop 的区别（唯一的两处）：
#   ① 拉数用**经理工号** 18591910491（店长号看不到华阳城）；
#   ② 经理号导出的是**全公司**明细/销售分析，故管线最前面多一道
#      「本地按门店过滤」（filter_maoli_hyc.py / fetch_sales_analysis_hyc.py），
#      calc_data.py / write_xs_xml.py 保持与 shop 逐字一致、不改编。
#
# 步骤：
#   1) 拉全公司门店毛利明细（浏览器导出，经理号）
#   2) 过滤出「华为华阳城合作店」→ 本店专用明细
#   3) 拉华阳城销售分析（纯 HTTP，经理号，页定位 ≈100s）
#   4) 写底表「销售分析」sheet（XML 级原位替换）
#   5) run_pipeline：写底表XS → 复算 → 渠道挂账 → 生成 index.html → 推送
#
# 【2026-09-20 提速改造】销售分析拉取 612s → ~100s（省 84%）：
#   · 接口每页有 ~50s 固定开销，与返回行数几乎无关 → pageSize 5000→20000、并发 5→8
#   · 接口按门店分组排序，华阳城只落在第 11/12 页 → sa_page_map.json 缓存命中页，
#     日常只拉 [10,11,12] 三页（±1 页兜底），行数校验不过自动降级全量重扫
#
# 用法：
#   bash update_hyc.sh                 # 全流程 + 推送上线
#   bash update_hyc.sh --no-push       # 只到本地 index.html
#   bash update_hyc.sh --skip-fetch    # 跳过联网拉数（复用本地已有明细）
#   bash update_hyc.sh --skip-sa       # 只跳销售分析（渠道卡停留上次拉取日 ⚠️）
#   bash update_hyc.sh --full-scan     # 强制全量重扫并重建页映射
# ============================================================
set -u

HERE="$(cd "$(dirname "$0")" && pwd)"
PY="/Users/mac/.workbuddy/binaries/python/envs/default/bin/python"
# 🔴 playwright 装在托管 workspace：不设 NODE_PATH 时 1/5 步会静默失败
#    （Error: Cannot find module 'playwright'，日志里只看到一串 '}' 极易误判），
#    随后管线继续用**旧明细**出数 —— 历史事故根因。
#    定时入口 run_hyc_scheduled.sh 已设一份；这里内置，保证手动跑同样安全。
export NODE_PATH="${NODE_PATH:-/Users/mac/.workbuddy/binaries/node/workspace/node_modules}"
MGR_DL="/Users/mac/.local/share/TeleAgent/playwright-mcp/hyc_mgr"   # 经理号导出的全公司明细
MAOLI_HYC="/Users/mac/.local/share/TeleAgent/playwright-mcp/hyc/华阳城门店毛利明细表-华为终端.xlsx"
TASK_XLSX="/Users/mac/Desktop/华阳城销售/华阳城9月任务进度.xlsx"

NO_PUSH=""; SKIP_FETCH=""; FORCE=""; SKIP_SA=""; SA_ARGS=""
for a in "$@"; do
  case "$a" in
    --no-push) NO_PUSH="--no-push" ;;
    --skip-fetch) SKIP_FETCH=1 ;;
    --force) FORCE=1 ;;
    --skip-sa) SKIP_SA=1 ;;
    --full-scan) SA_ARGS="--no-locate" ;;
  esac
done

cd "$HERE" || exit 1

# ============================================================
# 互斥锁（2026-09-20 加）：定时档位密集时，防止两个实例并发跑同一份底表
#   背景：单次全流程约 12 分钟，而排期含 21:57 / 22:02 这种间隔仅 5 分钟的档位，
#         并发会同时写 sa_hyc_month.json / sa_aug_cache.json / 底表 → 数据互相覆盖。
#   机制：mkdir 原子锁（macOS 无 flock）。后启动者直接跳过，绝不并行。
#   手动强制跑：bash update_hyc.sh --force
# ============================================================
LOCK_DIR="/tmp/shophyc_update.lock.d"
if [ -z "$FORCE" ]; then
  if [ -d "$LOCK_DIR" ]; then
    # 僵尸锁保护：正常流程 <= 20 分钟，超 40 分钟视为进程已死
    if [ -n "$(find "$LOCK_DIR" -maxdepth 0 -mmin +40 2>/dev/null)" ]; then
      echo "⚠️  检测到超过 40 分钟的陈旧锁，强制释放"
      rm -f "$LOCK_DIR/pid" 2>/dev/null; rmdir "$LOCK_DIR" 2>/dev/null
    fi
  fi
  if ! mkdir "$LOCK_DIR" 2>/dev/null; then
    echo "⏭️  已有华阳城更新实例在跑（锁 ${LOCK_DIR}，pid $(cat "${LOCK_DIR}/pid" 2>/dev/null || echo '?')），"
    echo "    本次跳过，以免并发写坏底表。要强制跑请加 --force"
    exit 0
  fi
  echo "$$" > "$LOCK_DIR/pid"
  trap 'rm -f "$LOCK_DIR/pid" 2>/dev/null; rmdir "$LOCK_DIR" 2>/dev/null' EXIT INT TERM
  echo "🔒 已获取互斥锁（pid $$）"
else
  echo "⚠️  --force：跳过互斥锁检查（请确认没有其他实例在跑）"
fi

# ============================================================
# 登录态体检（2026-09-20 加）：防「经理号被李家村流程覆盖」
#   事故：2026-09-20 10:23 李家村流程把店长号登录态写进 yonyou-mgr-default.json，
#         导致华阳城拉数是「李家村单店」→ 华阳城 0 行 → 14:07/18:12 两档全失败。
#   判据：mgr 状态文件与店长号状态文件**内容完全相同** = 被覆盖。
#   处置：自动重登经理号（relogin_mgr.sh），再进主流程。
# ============================================================
MGR_STATE="$HOME/.agent-browser/sessions/yonyou-mgr-default.json"
LJC_STATE="$HOME/.agent-browser/sessions/yonyou-default.json"
if [ -z "$SKIP_FETCH" ] && [ -f "$MGR_STATE" ] && [ -f "$LJC_STATE" ]; then
  if cmp -s "$MGR_STATE" "$LJC_STATE"; then
    echo "⚠️  经理号登录态与店长号完全相同（疑似被李家村流程覆盖），自动重登…"
    bash "$HERE/relogin_mgr.sh" 2>&1 | tail -3
    if cmp -s "$MGR_STATE" "$LJC_STATE"; then
      echo "❌ 重登后仍与店长号相同，终止（避免用错账号出数）"; exit 1
    fi
    echo "✓ 经理号登录态已恢复"
  fi
fi

if [ -z "$SKIP_FETCH" ]; then
  echo "===== 1/5 拉全公司门店毛利明细（经理号） ====="
  PREV_MT="$(stat -f%m "$(ls -t "$MGR_DL"/*.xlsx 2>/dev/null | head -1)" 2>/dev/null || echo 0)"
  PROFIT_DL="$MGR_DL" /Users/mac/.workbuddy/binaries/node/versions/22.22.2-3/bin/node fetch_profit_mgr.cjs 2>&1 | tail -3
  SRC="$(ls -t "$MGR_DL"/*.xlsx 2>/dev/null | head -1)"
  if [ -z "$SRC" ]; then echo "❌ 未拿到毛利明细，终止"; exit 1; fi
  echo "  源文件: $SRC"
  # 🔴 新鲜度闸门：导出成功必然刷新文件 mtime。若 mtime 没变且已 >30 分钟，
  #    说明导出静默失败（历史事故：NODE_PATH 缺失 → 沿用旧明细出数，数字陈旧却无人察觉）。
  CUR_MT="$(stat -f%m "$SRC" 2>/dev/null || echo 0)"
  if [ "$CUR_MT" -le "$PREV_MT" ]; then
    AGE=$(( $(date +%s) - CUR_MT ))
    if [ "$AGE" -gt 1800 ]; then
      echo "❌ 毛利明细未刷新（mtime 仍停在 $((AGE/60)) 分钟前），疑似导出失败。"
      echo "   已中止，避免用旧明细出数；确认要复用旧数据请显式加 --skip-fetch。"
      exit 1
    fi
    echo "  ⚠️ 本次导出未产生新文件，沿用 ${AGE}s 前的明细"
  fi

  echo "===== 2/5 过滤出「华为华阳城合作店」 ====="
  "$PY" filter_maoli_hyc.py "$SRC" "$MAOLI_HYC" || exit 1

  if [ -z "$SKIP_SA" ]; then
    echo "===== 3/5 拉华阳城销售分析（纯 HTTP，经理号，页定位） ====="
    # 失败自愈（2026-09-20 加）：最常见根因 = 经理号登录态被别的流程覆盖
    #   （拉回的其实是别的店，"本月无华阳城记录" exit 1）→ 自动重登经理号重试一次
    if ! "$PY" fetch_sales_analysis_hyc.py $SA_ARGS; then
      echo "⚠️  销售分析拉取失败，自动重登经理号并重试一次…"
      bash "$HERE/relogin_mgr.sh" 2>&1 | tail -3
      "$PY" fetch_sales_analysis_hyc.py $SA_ARGS || exit 1
    fi

    echo "===== 4/5 写底表「销售分析」sheet ====="
    "$PY" update_sales_analysis.py sa_hyc_month.json "$TASK_XLSX" || exit 1
  else
    echo "===== 3-4/5 --skip-sa：跳过销售分析拉取 ⚠️ ====="
    [ -f sa_aug_cache.json ] || { echo "❌ sa_aug_cache.json 不存在，不能跳过"; exit 1; }
    "$PY" -c "
import json
recs = json.load(open('sa_aug_cache.json', encoding='utf-8')).get('records', [])
ds = sorted(str(r.get('dDate') or '')[:10] for r in recs if r.get('dDate'))
print('   ⚠️ 渠道卡复用缓存 %d 行，流水 %s ~ %s（非今日新数据，请自行判断是否可接受）'
      % (len(recs), ds[0] if ds else '—', ds[-1] if ds else '—'))
"
  fi
else
  echo "（--skip-fetch）复用本地明细: $MAOLI_HYC"
  [ -f "$MAOLI_HYC" ] || { echo "❌ 本地明细不存在，去掉 --skip-fetch"; exit 1; }
fi

echo "===== 5/5 跑管线（写XS → 复算 → 渠道 → 建页 → 推送） ====="
# --no-sa：sa_aug_cache.json 已由 fetch_sales_analysis_hyc.py 以本店口径写好，
#          不能让 update_sa_cache.py 用 sa_xlsx 覆盖。
"$PY" run_pipeline.py "$MAOLI_HYC" sa_hyc_month.json --no-sa $NO_PUSH
RC=$?

echo
echo "===== 校验（独立复算底表 SUMIFS ↔ data.json 逐格对照） ====="
"$PY" verify_hyc.py "$MAOLI_HYC" data.json

exit $RC
