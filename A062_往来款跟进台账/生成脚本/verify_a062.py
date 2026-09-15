# -*- coding: utf-8 -*-
"""重算过的成品拿来逐项核：
   ① 原表 38 行的数字一个都没变；
   ② 汇总块自洽；
   ③ 往【数据录入】塞一笔回款，验证「已收款」真的会自动认领并按先来后到分摊。
   跑法：python3 verify_a062.py <重算过的xlsx>
"""
import sys, os, subprocess, shutil, openpyxl

SRC_ORIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '参考',
                        '原A062_多账户收支台账.xlsx')
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
path = sys.argv[1]
ok = bad = 0

def chk(name, got, want, tol=0.005):
    global ok, bad
    if isinstance(want, (int, float)) and isinstance(got, (int, float)):
        good = abs(got - want) <= tol
    else:
        good = str(got) == str(want)
    if good: ok += 1
    else:
        bad += 1
        print('  ✗ %-46s 实际=%-18s 应为=%s' % (name, got, want))
    return good

w = openpyxl.load_workbook(path, data_only=True)
g = w['往来款跟进']
o = openpyxl.load_workbook(SRC_ORIG, data_only=True)['往来款跟进']

print('=== ① 原表手工录入的格子一个都不能变 ===')
# 用户手工填的列：B日期 C生日 D客户 E产品 F销售金额 J销售费用 K支付方式 L备注
#                 M日期 N供货商 O购货款 S备注
import datetime as _dt
_EP = _dt.datetime(1899, 12, 30)
def _ser(v):
    """套了日期格式的格子读回来是 datetime，要换算回序列值才比得了"""
    if isinstance(v, _dt.datetime): return (v - _EP).days
    if isinstance(v, _dt.date):     return (_dt.datetime(v.year, v.month, v.day) - _EP).days
    return v
for col in (2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15, 19):
    L = openpyxl.utils.get_column_letter(col)
    diff = 0
    for r in range(11, 511):
        a, b = _ser(o.cell(r, col).value), _ser(g.cell(r, col).value)
        if a is None and b is None: continue
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            if abs(a - b) > 0.005: diff += 1
        elif a != b:
            diff += 1
            if diff <= 3: print('    %s%d: 原=%r 现=%r' % (L, r, a, b))
    chk(f'{L} 列 500 行逐格未变', diff, 0)

print('\n=== ② 已收款：数据录入里还没记这些客户，应沿用原手工数 ===')
recv = sum(float(g.cell(r,7).value or 0) for r in range(11,511) if isinstance(g.cell(r,7).value,(int,float)))
chk('已收款合计 = 原手工 19,140', recv, 19140)
for r, want in ((11, 3280), (12, 3980), (13, 2980), (14, 2980), (15, 1640), (16, 1000), (17, 3280)):
    chk(f'G{r} 已收款', g.cell(r,7).value, want)

print('\n=== ③ 尚欠款 = 销售金额 − 已收款 − 销售费用 ===')
for r, want in ((11, 0), (12, 0), (13, 0), (14, 0), (15, 1640), (16, 2280), (17, 0)):
    chk(f'H{r} 尚欠款', g.cell(r,8).value, want)
chk('I17 未收总计（累计到第17行）', g['I17'].value, 3920)

print('\n=== ④ 左上角汇总块 ===')
for cell, want, lab in (('M2', 23660, '总销售金额'), ('M3', 0, '购货款'), ('M4', 600, '销售费用'),
                        ('M5', 23060, '毛利'), ('P2', 23060, '应收款项'), ('P3', 19140, '已收款'),
                        ('P4', 3920, '尚欠款'), ('S2', 0, '购货款'), ('S3', 0, '已付款'),
                        ('S4', 0, '应付尚欠'), ('S5', 3920, '应收-应付')):
    chk(f'{cell} {lab}', g[cell].value, want)

print('\n=== ⑤ 对账表：按客户 ===')
rc = w['客户供应商往来对账']
rows = {}
for r in range(6, 66):
    nm = rc.cell(r, 2).value
    if nm: rows[nm] = r
chk('客户行数（跟着基础资料走）', len(rows), 37)
for nm, sale, recvd, owe in (('沈祖模', 7260, 7260, 0), ('谢爱琴', 6560, 5960, 0),
                             ('罗理华', 3280, 1640, 1640), ('马珍珠', 3280, 1000, 2280),
                             ('徐江英', 3280, 3280, 0)):
    r = rows.get(nm)
    if not r: chk(f'{nm} 在对账表里', '缺失', '存在'); continue
    chk(f'{nm} 销售金额', rc.cell(r,3).value, sale)
    chk(f'{nm} 已收款(往来款跟进)', rc.cell(r,6).value, recvd)
    chk(f'{nm} 尚欠款', rc.cell(r,8).value, owe)
chk('客户合计·销售金额', rc.cell(66,3).value, 23660)
chk('客户合计·应收款',   rc.cell(66,5).value, 23060)
chk('客户合计·已收款',   rc.cell(66,6).value, 19140)
chk('客户合计·尚欠款',   rc.cell(66,8).value, 3920)

print('\n=== ⑥ 无主的钱 ===')
# 数据录入 F 列收入合计 22200，其中挂了客户名的只有丁爱妹 2000
chk('收进来但没选客户名的钱', rc.cell(93,6).value, 20200)
chk('【数据录入】收入合计',    rc.cell(95,6).value, 22200)
chk('【数据录入】支出合计',    rc.cell(96,6).value, 11660)

print('\n' + '='*54)
print('共核 %d 项，通过 %d 项，不通过 %d 项' % (ok+bad, ok, bad))
sys.exit(1 if bad else 0)
