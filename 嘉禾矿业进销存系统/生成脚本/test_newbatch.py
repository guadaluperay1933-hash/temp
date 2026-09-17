# -*- coding: utf-8 -*-
"""验一件事：在【入库单】/【费用台账】录一个**新批次**，【到货批次】会不会自己长出来。

用法: python3 test_newbatch.py <交付件xlsx> <工作目录>
往入库单塞 2 行新柜、往费用台账塞 1 笔「费用先到、货还没到」的批次，
重算完看到货批次是不是自动多了 2 行、柜号日期货源对不对。
"""
import os, sys, shutil, datetime, subprocess
import openpyxl

SRC, WORK, BASE = sys.argv[1], sys.argv[2], sys.argv[3]
D0, S0, IN_END, F_END, B_END = 4, 5, 4003, 803, 204
NEW_BATCH = 'TESTU7654321-202609'
FEE_BATCH = 'FEEFIRST-202610'

os.makedirs(WORK, exist_ok=True)
tgt = os.path.join(WORK, 'newbatch.xlsx')
shutil.copy(SRC, tgt)
wb = openpyxl.load_workbook(tgt)
inb, fee, bat = wb['入库单'], wb['费用台账'], wb['到货批次']

# 基准清单必须从**重算过**的文件读：交付件里 A 列是公式，读出来是公式原文不是批次号
before = [c.value for c in
          openpyxl.load_workbook(BASE, data_only=True)['到货批次']['A'][S0 - 1:B_END]]
last = max(r for r in range(D0, IN_END + 1) if inb.cell(r, 5).value) + 1
code = next(inb.cell(r, 7).value for r in range(D0, IN_END + 1) if inb.cell(r, 7).value)
for k, (qty, price) in enumerate(((10, 25.5), (4, 100))):
    r = last + k
    for c, v in ((3, datetime.datetime(2026, 9, 8 + k)), (4, 'A国内'), (5, NEW_BATCH),
                 (6, '测试供应商'), (7, code), (11, qty), (12, price),
                 (24, '采购入库'), (25, '正常'), (26, '正常'), (27, '测试')):
        r_ = inb.cell(r, c); r_.value = v
lastf = max(r for r in range(D0, F_END + 1) if fee.cell(r, 3).value) + 1
for c, v in ((2, datetime.datetime(2026, 10, 3)), (3, FEE_BATCH), (4, '国内海运费'),
             (5, '测试货代'), (6, 'CNY'), (7, 12345), (10, '是'), (11, '正常'), (12, '费用先到')):
    fee.cell(lastf, c).value = v
# 再给新柜加一笔运费，看运费率和摊平状态
for c, v in ((2, datetime.datetime(2026, 9, 10)), (3, NEW_BATCH), (4, '国内海运费'),
             (5, '测试货代'), (6, 'CNY'), (7, 500), (10, '是'), (11, '正常'), (12, '测试运费')):
    fee.cell(lastf + 1, c).value = v
wb.save(tgt)
print(f'塞进去：入库单 2 行（批次 {NEW_BATCH}）、费用台账 2 笔'
      f'（{NEW_BATCH} 运费 500 元、{FEE_BATCH} 只有费用没有货）')

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), '工具'))
import recalc_safe                                                   # noqa: E402
recalc_safe.recalc(tgt, 900)

got = openpyxl.load_workbook(tgt, data_only=True)['到货批次']
after = [got.cell(r, 1).value for r in range(S0, B_END + 1)]
new = [x for x in after if x and x not in before]
ok = bad = 0


def chk(n, g, w):
    global ok, bad
    if g == w:
        ok += 1
    else:
        bad += 1
        print(f'  ✗ {n}: 实际 {g!r}  应为 {w!r}')


print('原来 %d 个批次 → 现在 %d 个；新长出来的：%s'
      % (len([x for x in before if x]), len([x for x in after if x]), new))
chk('新长出 2 个批次', sorted(new), sorted([NEW_BATCH, FEE_BATCH]))
row = {got.cell(r, 1).value: r for r in range(S0, B_END + 1) if got.cell(r, 1).value}
r1 = row.get(NEW_BATCH)
if r1:
    v = [got.cell(r1, c).value for c in range(1, 14)]
    print('  新柜那一行：', [str(x)[:19] for x in v])
    chk('柜号从批次号截出来', v[1], 'TESTU7654321')
    chk('到港日期 = 第一笔入库日', v[2], datetime.datetime(2026, 9, 8))
    chk('货源自动带出', v[3], 'A国内')
    chk('物料行数', v[4], 2)
    chk('货值·先令 = (10×25.5+4×100)×370', round(v[6]), round((10 * 25.5 + 4 * 100) * 370))
    chk('费用·先令 = 500×370', round(v[7]), 500 * 370)
    chk('已摊运费 = 费用（没有待补价行，整批摊平）', round(v[9]), 500 * 370)
    chk('未摊挂账 = 0', round(v[10]), 0)
    chk('状态', v[12], '已摊平')
r2 = row.get(FEE_BATCH)
if r2:
    v = [got.cell(r2, c).value for c in range(1, 14)]
    print('  只有费用那一行：', [str(x)[:19] for x in v])
    chk('柜号', v[1], 'FEEFIRST')
    chk('到港日期退而取费用日期', v[2], datetime.datetime(2026, 10, 3))
    chk('物料行数 0', v[4], 0)
    chk('费用·先令 = 12345×370', round(v[7]), 12345 * 370)
    chk('状态点名「只有费用没有货」', v[12], '△这个批次只有费用、还没有入库单（货还没到？）')
print(f'\n══ 新批次自动生成：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
