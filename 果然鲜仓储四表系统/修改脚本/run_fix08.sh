#!/bin/bash
# A050 第八轮 —— 从「参考/本轮」的四个原件，一路跑到可以交付的四个文件
set -e
ROOT=/home/user/temp/果然鲜仓储四表系统
TOOLS=/home/user/temp/工具
SRC=$ROOT/参考/本轮
OUT=${1:-/tmp/claude-0/deliver}
rm -rf "$OUT"; mkdir -p "$OUT"

declare -A NAME=( [01]=01_水果进销存台账模板 [02]=02_物料与周转物台账模板
                  [03]=03_财务账套与报表模板 [04]=04_综合查询对账单模板 )

for n in 01 02 03 04; do
  echo "══════ $n ══════"
  python3 "$ROOT/修改脚本/fix08_$n.py" "$SRC/$n.xlsx" "$OUT/${NAME[$n]}.xlsx"
done

echo "══════ 补回跨文件链接关系 ══════"
for n in 03 04; do python3 "$ROOT/修改脚本/repair_ext_links.py" "$OUT/${NAME[$n]}.xlsx" "$SRC/$n.xlsx"; done

echo "══════ 括号/引号自检 ══════"
for n in 01 02 03 04; do python3 "$TOOLS/check_formula.py" "$OUT/${NAME[$n]}.xlsx"; done

echo "══════ 工作组自检 ══════"
for n in 01 02 03 04; do python3 "$TOOLS/fix_sheet_selection.py" "$OUT/${NAME[$n]}.xlsx" || true; done

cd "$OUT"
echo "══════ 重算 01 / 02 ══════"
python3 "$TOOLS/recalc_safe.py" "${NAME[01]}.xlsx" 900 || true
python3 "$TOOLS/recalc_safe.py" "${NAME[02]}.xlsx" 900 || true

echo "══════ 刷新 03 / 04 的跨文件缓存并重算 ══════"
for n in 03 04; do
  python3 "$ROOT/修改脚本/refresh_link_cache.py" "${NAME[$n]}.xlsx" "${NAME[01]}.xlsx" "${NAME[02]}.xlsx"
  python3 "$TOOLS/recalc_safe.py" "${NAME[$n]}.xlsx" 900 || true
done

echo "══════ 重算后再查一次工作组 ══════"
for n in 01 02 03 04; do python3 "$TOOLS/fix_sheet_selection.py" "${NAME[$n]}.xlsx" || true; done
ls -la "$OUT"
