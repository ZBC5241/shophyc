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
#   3) 拉华阳城销售分析（纯 HTTP，经理号，46 页全量）
#   4) 写底表「销售分析」sheet（XML 级原位替换）
#   5) run_pipeline：写底表XS → 复算 → 渠道挂账 → 生成 index.html → 推送
#
# 用法：
#   bash update_hyc.sh                 # 全流程 + 推送上线
#   bash update_hyc.sh --no-push       # 只到本地 index.html
#   bash update_hyc.sh --skip-fetch    # 跳过联网拉数（复用本地已有明细）
# ============================================================
set -u

HERE="$(cd "$(dirname "$0")" && pwd)"
PY="/Users/mac/.workbuddy/binaries/python/envs/default/bin/python"
MGR_DL="/Users/mac/.local/share/TeleAgent/playwright-mcp/hyc_mgr"   # 经理号导出的全公司明细
MAOLI_HYC="/Users/mac/.local/share/TeleAgent/playwright-mcp/hyc/华阳城门店毛利明细表-华为终端.xlsx"
TASK_XLSX="/Users/mac/Desktop/华阳城销售/华阳城9月任务进度.xlsx"

NO_PUSH=""; SKIP_FETCH=""
for a in "$@"; do
  case "$a" in
    --no-push) NO_PUSH="--no-push" ;;
    --skip-fetch) SKIP_FETCH=1 ;;
  esac
done

cd "$HERE" || exit 1

if [ -z "$SKIP_FETCH" ]; then
  echo "===== 1/5 拉全公司门店毛利明细（经理号） ====="
  PROFIT_DL="$MGR_DL" /Users/mac/.workbuddy/binaries/node/versions/22.22.2-3/bin/node fetch_profit_mgr.cjs 2>&1 | tail -3
  SRC="$(ls -t "$MGR_DL"/*.xlsx 2>/dev/null | head -1)"
  if [ -z "$SRC" ]; then echo "❌ 未拿到毛利明细，终止"; exit 1; fi
  echo "  源文件: $SRC"

  echo "===== 2/5 过滤出「华为华阳城合作店」 ====="
  "$PY" filter_maoli_hyc.py "$SRC" "$MAOLI_HYC" || exit 1

  echo "===== 3/5 拉华阳城销售分析（纯 HTTP，经理号） ====="
  "$PY" fetch_sales_analysis_hyc.py || exit 1

  echo "===== 4/5 写底表「销售分析」sheet ====="
  "$PY" update_sales_analysis.py sa_hyc_month.json "$TASK_XLSX" || exit 1
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
