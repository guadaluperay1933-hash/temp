# -*- coding: utf-8 -*-
"""换参数跑一遍：月份 / 公司 / 日期区间 / 项目 / 含税开关，看勾稽校验是不是永远成立。"""
import sys, json, subprocess, openpyxl, datetime
BASE='A052_final2.xlsx'
def run(tag, mutate):
    f=f'sc_{tag}.xlsx'
    wb=openpyxl.load_workbook(BASE); mutate(wb); wb.save(f)
    out=subprocess.run(['python3','/mnt/skills/public/xlsx/scripts/recalc.py',f],
                       capture_output=True,text=True,timeout=1800).stdout
    d=json.loads(out[out.find('{'):])
    v=openpyxl.load_workbook(f,data_only=True); R=v['月度汇报表']
    bad=[]
    for r in range(174,184):
        if R.cell(r,5).value and str(R.cell(r,5).value).startswith('✗'):
            bad.append(f'{R.cell(r,1).value}: {R.cell(r,5).value}')
    print(f'【{tag}】重算 {d.get("status")} 错误 {d.get("total_errors")} | 校验 {"全部√" if not bad else "✗ "+ "; ".join(bad)}')
    return v

SC = [
 ('月份202601', lambda wb: wb['月度汇报表'].__setitem__('B2', 202601)),
 ('月份202604', lambda wb: wb['月度汇报表'].__setitem__('B2', 202604)),
 ('公司瑞盛天', lambda wb: wb['月度汇报表'].__setitem__('E2', '瑞盛天')),
 ('公司卓卓通', lambda wb: wb['月度汇报表'].__setitem__('E2', '卓卓通')),
 ('月份+公司',  lambda wb: (wb['月度汇报表'].__setitem__('B2', 202603), wb['月度汇报表'].__setitem__('E2','双思威'))),
 ('无数据月份', lambda wb: wb['月度汇报表'].__setitem__('B2', 202612)),
 ('不含税口径', lambda wb: wb['基础资料'].__setitem__('AH4','不含税金额')),
]
res={}
for tag,mut in SC:
    res[tag]=run(tag,mut)
