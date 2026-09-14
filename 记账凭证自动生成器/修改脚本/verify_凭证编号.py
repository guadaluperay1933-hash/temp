# -*- coding: utf-8 -*-
"""改完之后逐张凭证核对：右上角有没有留空、辅助列的号对不对、开关能不能切回去"""
import sys, re
from openpyxl import load_workbook

XL = sys.argv[1] if len(sys.argv) > 1 else '../记账凭证自动生成器_2026.9.9.xlsx'
BLOCK = 15
SHEETS = [('凭证收入', '助手收入'), ('凭证回款', '助手回款'), ('凭证报销', '助手报销')]
wb = load_workbook(XL, data_only=True)
wf = load_workbook(XL, data_only=False)
ok = bad = 0
def chk(name, cond, detail=''):
    global ok, bad
    if cond: ok += 1
    else: bad += 1; print(f'  ✗ {name}  {detail}')

sw = wb['参数']['B9'].value
print(f'参数!B9「凭证号是否打印」= {sw!r}')
chk('开关默认为「否」', sw == '否', f'实际 {sw!r}')

for vn, hn in SHEETS:
    ws, wsf, hs = wb[vn], wf[vn], wb[hn]
    n_live = n_blank = n_num = 0
    for b in range(1, ws.max_row + 1, BLOCK):
        k = (b - 1) // BLOCK + 2
        hdr = ws.cell(row=b + 2, column=6).value
        m = ws.cell(row=b + 2, column=13).value
        n = ws.cell(row=b + 2, column=14).value
        o = ws.cell(row=b + 2, column=15).value
        if hs.cell(row=k, column=18).value in (None, ''):          # R 列为空 = 这页没凭证
            chk(f'{vn} 第{b}块 空页应无内容', (hdr in (None, '')) and (n in (None, '')),
                f'hdr={hdr!r} n={n!r}')
            continue
        n_live += 1
        chk(f'{vn} 第{b}块 右上角不能再有号',
            isinstance(hdr, str) and not re.search(r'字第\s*\d', hdr), f'{hdr!r}')
        if isinstance(hdr, str) and '字第' in hdr and re.search(r'字第[\s　]+号', hdr):
            n_blank += 1
        rk = hs.cell(row=k, column=18).value                        # R 列 序号＝第几家店/第几张单
        want_no = hs.cell(row=int(rk) + 1, column=16).value         # 该店那一行的 P 列 凭证号
        want_pg = hs.cell(row=k, column=17).value                   # Q 列 页码＝打印第几页
        chk(f'{vn} 第{b}块 N列凭证号', n == want_no, f'表内={n!r} 助手={want_no!r}')
        chk(f'{vn} 第{b}块 M列打印页', m == want_pg, f'表内={m!r} 助手={want_pg!r}')
        chk(f'{vn} 第{b}块 O列含该号', isinstance(o, str) and f'第 {want_no} 号' in o, f'{o!r}')
        if isinstance(n, int): n_num += 1
    print(f'{vn}: 有效凭证 {n_live} 张，右上角留空 {n_blank} 张，辅助列出号 {n_num} 张')
    chk(f'{vn} 全部有效凭证都留空', n_blank == n_live, f'{n_blank}/{n_live}')
    chk(f'{vn} 全部有效凭证辅助列都有号', n_num == n_live, f'{n_num}/{n_live}')
    chk(f'{vn} 打印区域固定为 B:I',
        str(wsf.print_area or '').endswith(f'$B$1:$I${ws.max_row}'), repr(wsf.print_area))
    chk(f'{vn} 分页符没动', len(wsf.row_breaks.brk) == ws.max_row // BLOCK, len(wsf.row_breaks.brk))

for hn in ('助手收入', '助手回款', '助手报销'):
    chk(f'{hn} 保持隐藏', wf[hn].sheet_state == 'hidden', wf[hn].sheet_state)

print(f'\n=========  {ok} 项通过，{bad} 项不通过  =========')
sys.exit(1 if bad else 0)
