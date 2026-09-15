# -*- coding: utf-8 -*-
"""行为测试：往【数据录入】里补几笔回款/付款，看【往来款跟进】会不会真的自动认领，
   而且按「先买的先冲」把钱分摊到同一个客户的多行上。
   跑法：python3 test_fifo.py <成品xlsx>
"""
import sys, os, shutil, subprocess, json, openpyxl

RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
SCR = '/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad'
src = sys.argv[1]
tmp = os.path.join(SCR, 'a062_fifo_test.xlsx')
shutil.rmtree(os.path.join(SCR, '.lo'), ignore_errors=True)
for junk in ('.~lock.a062_fifo_test.xlsx#',):
    try: os.remove(os.path.join(SCR, junk))
    except OSError: pass
shutil.copy(src, tmp)

wb = openpyxl.load_workbook(tmp)
sj = wb['数据录入']
import datetime as dt
# 在第一段空白流水行上补数据（原表流水到第 41 行为止）
ADD = [
    # 日期,        账户,   类别,   摘要,          收入,   支出, 产品, 客户,    供应商
    (dt.date(2026,9,20), '支付宝', '',       '沈祖模付货款', 5000, None, '羊奶', '沈祖模', None),
    (dt.date(2026,9,21), '微信',   '',       '罗理华付货款', 2000, None, '羊奶', '罗理华', None),
    (dt.date(2026,9,22), '农商',   '供应商', '付羊奶货款',   None, 3000, None,  None,    '羊奶供货商'),
]
for i, (d, acct, cat, memo, inc, out, prod, cust, supp) in enumerate(ADD):
    r = 42 + i
    sj.cell(r, 2).value = d;    sj.cell(r, 3).value = acct
    sj.cell(r, 4).value = cat;  sj.cell(r, 5).value = memo
    sj.cell(r, 6).value = inc;  sj.cell(r, 7).value = out
    sj.cell(r, 9).value = prod; sj.cell(r,10).value = cust
    sj.cell(r,11).value = supp
# 往来款跟进补一笔购货，才能验证应付侧的分摊
g = wb['往来款跟进']
g['M18'] = dt.date(2026,9,15); g['N18'] = '羊奶供货商'; g['O18'] = 1200
g['M19'] = dt.date(2026,9,18); g['N19'] = '羊奶供货商'; g['O19'] = 2500
wb.save(tmp)

print('重算中…')
r = subprocess.run([sys.executable, RECALC, tmp, '900', '--force'], capture_output=True, text=True)
print(r.stdout.strip()[:300])
# 原表本来就有 130 处 #NAME?（5 张账户分表用了 Excel 365 的 FILTER()，
# LibreOffice 不认）。只要不多出新的报错就算过。
BASE = 130
try:
    st = json.loads(r.stdout)
    assert st['total_errors'] <= BASE, st
    print('重算报错 %d 处（原表基线 %d），没多出新的' % (st['total_errors'], BASE))
except Exception as e:
    print('重算有问题:', e); sys.exit(1)

v = openpyxl.load_workbook(tmp, data_only=True)['往来款跟进']
ok = bad = 0
def chk(n, got, want, tol=0.005):
    global ok, bad
    good = abs((got or 0) - want) <= tol if isinstance(want,(int,float)) else got == want
    if good: ok += 1
    else: bad += 1; print('  ✗ %-42s 实际=%-14s 应为=%s' % (n, got, want))

print('\n=== 沈祖模：两行（3280 + 3980 = 7260），只付了 5000，应先冲第一行 ===')
chk('G11 已收款（第一行吃满 3280）', v['G11'].value, 3280)
chk('G12 已收款（第二行只剩 1720）', v['G12'].value, 1720)
chk('H11 尚欠款', v['H11'].value, 0)
chk('H12 尚欠款（3980-1720）',      v['H12'].value, 2260)
chk('两行已收合计 = 实付 5000', (v['G11'].value or 0)+(v['G12'].value or 0), 5000)

print('\n=== 罗理华：一行 3280，付了 2000 ===')
chk('G15 已收款', v['G15'].value, 2000)
chk('H15 尚欠款', v['H15'].value, 1280)

print('\n=== 谢爱琴：数据录入里没记，应继续沿用原手工数 ===')
chk('G13 已收款（原手工 2980）', v['G13'].value, 2980)
chk('G14 已收款（原手工 2980）', v['G14'].value, 2980)

# 注意：【数据录入】第 24 行原本就有一笔付给羊奶供货商的 5000，
# 加上这次补的 3000，一共付了 8000，而购货才 1200+2500=3700 —— 是付超了。
# MIN() 会把认领额卡在购货额上，超付的 4300 不会硬塞进来，
# 差额由【客户供应商往来对账】那张表的「差额」列兜出来。
print('\n=== 羊奶供货商：购货 1200+2500=3700，累计付了 5000+3000=8000（付超）===')
chk('P18 已付款（吃满本行购货 1200）', v['P18'].value, 1200)
chk('P19 已付款（吃满本行购货 2500）', v['P19'].value, 2500)
chk('Q18 尚欠款', v['Q18'].value, 0)
chk('Q19 尚欠款（付超了，不倒欠）',   v['Q19'].value, 0)
chk('R19 尚欠款总计',                v['R19'].value, 0)

print('\n=== 汇总块跟着动 ===')
chk('P3 已收款合计', v['P3'].value, 3280+1720+2980+2980+2000+1000+3280)
chk('S2 购货款合计', v['S2'].value, 3700)
chk('S3 已付款合计（认领额，封顶在购货额）', v['S3'].value, 3700)
chk('S4 应付尚欠',   v['S4'].value, 0)

print('\n=== 对账表：付超的部分要能看出来 ===')
rc = openpyxl.load_workbook(tmp, data_only=True)['客户供应商往来对账']
row = next((r for r in range(70, 90) if rc.cell(r, 2).value == '羊奶供货商'), None)
chk('羊奶供货商 在对账表里', row is not None, True)
if row:
    chk('购货款',                rc.cell(row, 3).value, 3700)
    chk('已付款(往来款跟进 认领)', rc.cell(row, 6).value, 3700)
    chk('已付款(数据录入 实付)',   rc.cell(row, 7).value, 8000)
    chk('差额 = 认领 − 实付（负数＝付超）', rc.cell(row, 9).value, -4300)

print('\n共核 %d 项，通过 %d，不通过 %d' % (ok+bad, ok, bad))
sys.exit(1 if bad else 0)
