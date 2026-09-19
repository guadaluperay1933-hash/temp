#!/bin/bash
# A050 第九轮 —— 在第八轮交付件的基础上打补丁
set -e
ROOT=/home/user/temp/果然鲜仓储四表系统
TOOLS=/home/user/temp/工具
OUT=${1:-/tmp/claude-0/deliver9}
rm -rf "$OUT"; mkdir -p "$OUT"

declare -A NAME=( [01]=01_水果进销存台账模板 [02]=02_物料与周转物台账模板
                  [03]=03_财务账套与报表模板 [04]=04_综合查询对账单模板 )

cp "$ROOT/${NAME[01]}.xlsx" "$OUT/"
cp "$ROOT/${NAME[03]}.xlsx" "$OUT/"
python3 "$ROOT/修改脚本/fix09_02.py" "$ROOT/${NAME[02]}.xlsx" "$OUT/${NAME[02]}.xlsx"
python3 "$ROOT/修改脚本/fix09_04.py" "$ROOT/${NAME[04]}.xlsx" "$OUT/${NAME[04]}.xlsx"

echo "══════ 补回跨文件链接关系 ══════"
python3 "$ROOT/修改脚本/repair_ext_links.py" "$OUT/${NAME[04]}.xlsx" "$ROOT/${NAME[04]}.xlsx"

echo "══════ 括号/引号自检 ══════"
for n in 01 02 03 04; do python3 "$TOOLS/check_formula.py" "$OUT/${NAME[$n]}.xlsx"; done

cd "$OUT"
echo "══════ 重算 02 ══════"
python3 "$TOOLS/recalc_safe.py" "${NAME[02]}.xlsx" 900 || true
echo "══════ 刷新 03 / 04 缓存并重算 ══════"
for n in 03 04; do
  python3 "$ROOT/修改脚本/refresh_link_cache.py" "${NAME[$n]}.xlsx" "${NAME[01]}.xlsx" "${NAME[02]}.xlsx"
  python3 "$TOOLS/recalc_safe.py" "${NAME[$n]}.xlsx" 900 || true
done
echo "══════ 工作组自检 ══════"
for n in 01 02 03 04; do python3 "$TOOLS/fix_sheet_selection.py" "${NAME[$n]}.xlsx" || true; done
ls -la "$OUT"
