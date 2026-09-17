# -*- coding: utf-8 -*-
"""把后补的三份文件并进 mine_data.json：

m1 原_物资流水帐.xlsx           2677 行数量流水（采购 / 借调 / 使用），7 个矿区
m2 原_海运费用统计_FJMZL-JH2026013  一个新货柜（MSKU9402161）的采购清单 + 国内段费用
m3 原_达市到货费用统计表.xlsx      27 个柜的坦桑当地到港费用（TBS/船公司/港口/GALCO/滞箱/堆场/柜运费/海关税/免税/放行费）

两处要当心的地方：
1. m3 的「船公司费用」列本该填美金（实际值 128~545），但 EMCU1461128 和 MSKV4646463
   两行填成了先令（1,351,352.5 和 232,910.41）。按美金乘 2650 会多出 42 亿先令。
   规则：这一列大于 5000 的按先令处理 —— 这样算出来跟他们自己的「总运费」列分毫不差。
2. 流水帐的「采购」和货柜到货**对不上**：93 个品种两边都有，只有 10 种数量一样
   （半胶手套 货柜2088 / 流水3751，坑道服 货柜200 / 流水730…）。
   两份都导进来，重叠的品种在入库单标出来，让用户定以哪份为准 —— 不偷偷丢也不偷偷加。
"""
import json, os, re, datetime, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
REF = os.path.join(ROOT, '参考')
DJ = os.path.join(HERE, 'mine_data.json')
D = json.load(open(DJ, encoding='utf-8'))

N = lambda v: float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
T = lambda v: ('' if v is None else str(v)).strip()
norm = lambda s: re.sub(r'[\s\-－—×xX*＊·．.、/／（）()]+', '', str(s)).upper()

def d2s(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.date):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime('%Y-%m-%d')
    return ''

# 现有物料：按规范化品名建索引，重名的沿用原编码，不另开新码
code_of = {}
for i in D['items']:
    code_of.setdefault(norm(i['name']), i['code'])
next_no = len(D['items']) + 1

def get_code(name, unit, cat_hint=''):
    global next_no
    k = norm(name)
    if k in code_of:
        return code_of[k], False
    code = 'JH%04d' % next_no
    next_no += 1
    code_of[k] = code
    D['items'].append(dict(code=code, name=name, spec='', unit=unit or '个',
                           cat=classify(name, '', cat_hint),
                           key='是' if classify(name, '', cat_hint) in KEY else '否'))
    return code, True

CLASS = [
    (('炸药', '雷管', '导爆', '乳化', '硝铵', '起爆', '导火', '电管'), '炸药火工'),
    (('柴油', '汽油', '机油', '液压油', '齿轮油', '黄油', '润滑', '防冻液', '加油', '刹车油'), '油料'),
    (('黄原酸', '硫酸铜', '碳酸钠', '松醇油', '硅酸钠', '硫化钠', 'JB-05', '二丁基',
      '药剂', '起泡剂', '捕收剂', '水玻璃', '石灰'), '选矿药剂'),
    (('钢球', '钢段', '钢锻'), '磨矿介质'),
    (('钢板', '花纹板', '焊管', '无缝管', '角钢', '扁钢', '钢轨', '轻轨', '圆钢', '槽钢',
      '工字钢', '方管', '钢管', '镀锌', '钢材', '型钢', '铁丝', '钢丝绳'), '钢材'),
    (('电缆', '电线', '屏蔽线', '开关', '启动柜', '变压器', '配电', '电柜', '控制柜',
      '接触器', '断路器', '电机', '灯', '铜鼻子', '铜铝鼻子', '空开', '电池'), '电缆电气'),
    (('球磨机', '破碎机', '浮选机', '分级机', '绞车', '空压机', '凿岩机', '水泵', '风机',
      '通风机', '叉车', '地磅', '发电机', '皮带秤', '压滤机', '给矿机', '振动筛',
      '搅拌', '装载机', '汽车衡', '切割机', '电焊机', '钻机'), '机械设备'),
    (('衬板', '叶轮', '轴承', '皮带', '滤芯', '密封', '刮板', '齿圈', '联轴器', '稳钉',
      '护板', '破碎壁', '轧臼壁', '扎臼壁', '刹车带', '配件', '备件', '钎子', '钻杆'), '机械配件'),
    (('螺栓', '螺母', '螺丝', '垫片', '金属垫', '胶垫', '球阀', '闸阀', '法兰', '对丝',
      '弯头', '三通', '软连接', '阀', '管件', '扳手', '钳', '锤', '锯', '钻头', '焊条',
      '水管', '丝牙', '葫芦', '榔头', '卷尺'), '五金管件'),
    (('手套', '工作服', '安全帽', '雨鞋', '口罩', '劳保', '棉被', '毛巾', '雨衣', '坑道服'), '劳保用品'),
    (('集装箱',), '集装箱'),
    (('土工膜', '吨袋', '编织袋', '篷布', '水泥'), '包装材料'),
]
KEY = {'炸药火工', '油料', '选矿药剂', '磨矿介质'}

def classify(name, spec, hint):
    s = f'{name} {spec} {hint}'
    for keys, c in CLASS:
        for k in keys:
            if k in s:
                return c
    return '其他'

# ══ m3：27 个柜的达市到港费用 ═══════════════════════════════════════
ws = openpyxl.load_workbook(os.path.join(REF, '原_达市到货费用统计表.xlsx'), data_only=True)['货柜为准校对']
COST = [('M', 'TBS费用', False), ('O', '船公司费用', True), ('Q', '港口费', False),
        ('S', 'GALCO海关操作费', False), ('T', '滞箱延误费', True), ('V', '堆场费', False),
        ('X', '集装箱运费', False), ('Z', '海关税', False), ('AB', '免税税费', False),
        ('AD', '放行费及代理费', False)]
want = {b['container']: b['batch'] for b in D['batches'] if b['container']}
local_fees, fixed = [], []
for r in range(3, 261):
    c = T(ws.cell(r, 8).value)
    if c not in want:
        continue
    al = N(ws.cell(r, 38).value) or 2650
    af = N(ws.cell(r, 32).value)
    got = 0
    for col, name, is_usd in COST:
        i = openpyxl.utils.column_index_from_string(col)
        v = N(ws.cell(r, i).value)
        if not v:
            continue
        if is_usd and v > 5000:          # 这一格填的是先令不是美金
            tzs = round(v)
            fixed.append((c, name, v))
        else:
            tzs = round(v * (al if is_usd else 1))
        got += tzs
        local_fees.append(dict(batch=want[c], container=c, date=d2s(ws.cell(r, 2).value),
                               cat=name, tzs=tzs, note=f'达市到港费用表第 {r} 行'))
    D_ = round(got - af)
    if abs(D_) > 2:
        print('  ⚠ %s 重算 %s 与表里「总运费」%s 差 %s' % (c, f'{got:,.0f}', f'{af:,.0f}', f'{D_:,.0f}'))

# ══ m2：新货柜 MSKU9402161 ═════════════════════════════════════════
wb2 = openpyxl.load_workbook(os.path.join(REF, '原_海运费用统计_FJMZL-JH2026013.xlsx'), data_only=True)
dt = wb2['明细表']
FEE_WORDS = ('海运费', '服务费', '搬运费', '叉车', '超时', '餐费', '装柜', '报关', '清关',
             '港杂', '拖车', '运费', '水费', '续单')
NEW_BATCH, NEW_DATE = 'MSKU9402161-202607', '2026-07-04'
new_in, new_fee = [], []
for r in range(4, 161):
    nm = T(dt.cell(r, 2).value)
    if not nm or nm == '合计':
        continue
    qty, price, amt = N(dt.cell(r, 6).value), N(dt.cell(r, 7).value), N(dt.cell(r, 8).value)
    if round(amt, 2) == 0:
        continue
    if any(w in nm for w in FEE_WORDS):
        new_fee.append(dict(batch=NEW_BATCH, date=NEW_DATE, cat=(
            '国内海运费' if '海运' in nm else '报关服务费' if ('报关' in nm or '续单' in nm)
            else '装柜费' if '装柜' in nm or '叉车' in nm or '搬运' in nm
            else '国内内陆运费'), cny=round(amt, 2),
            note=f'{nm}（FJMZL-JH2026013 / 提单 XMCX60772400，明细表第 {r} 行）'))
        continue
    if abs(round(qty * price, 2) - round(amt, 2)) > 0.004:
        qty, price = 1.0, round(amt, 2)
    code, _ = get_code(nm, T(dt.cell(r, 5).value))
    new_in.append(dict(date=NEW_DATE, supplier='福建马扎罗贸易有限公司', container='MSKU9402161',
                       batch=NEW_BATCH, source='A国内', code=code, name=nm,
                       spec=T(dt.cell(r, 4).value), unit_clean=T(dt.cell(r, 5).value) or '个',
                       qty=qty, price=price, amount=round(amt, 2), qout=0,
                       note=f'FJMZL-JH2026013 明细表第 {r} 行', row=r, hint=''))

# ══ m1：2677 行物资流水 ════════════════════════════════════════════
fl = openpyxl.load_workbook(os.path.join(REF, '原_物资流水帐.xlsx'), data_only=True)['出入库流水']
MINE_FIX = {'一矿': '一矿区', '二矿': '二矿区', '三矿': '三矿区', '六矿': '六矿区',
            '选厂': '选厂', '维修部': '维修部', '沙金': '沙金'}
buy, use, lend = [], [], []
n_new = 0
nodate = 0
for r in range(3, 2680):
    nm = T(fl.cell(r, 2).value)
    if not nm:
        continue
    unit = T(fl.cell(r, 3).value) or '个'
    kind = T(fl.cell(r, 4).value)
    qin, qout = N(fl.cell(r, 5).value), N(fl.cell(r, 6).value)
    to_, from_ = MINE_FIX.get(T(fl.cell(r, 8).value), ''), MINE_FIX.get(T(fl.cell(r, 9).value), '')
    d = d2s(fl.cell(r, 1).value)
    if not d:
        nodate += 1
    code, isnew = get_code(nm, unit)
    n_new += isnew
    base = dict(row=r, date=d, code=code, name=nm, unit=unit)
    if kind == '采购' and qin > 0:
        ym = d[:7].replace('-', '') if d else '000000'
        buy.append(dict(base, qty=qin, mine=to_, batch=f'本地-{ym}',
                        note=f'物资流水帐第 {r} 行 · {to_ or "未填矿区"}采购'))
    elif kind == '使用' and qout > 0:
        use.append(dict(base, qty=qout, mine=from_,
                        note=f'物资流水帐第 {r} 行'))
    elif kind == '借调' and qout > 0:
        lend.append(dict(base, qty=qout, to_mine=to_, from_mine=from_,
                         note=f'物资流水帐第 {r} 行 · {from_}→{to_}'))

# 重叠嫌疑：同一个物料，货柜里有、本地采购也有
cont_codes = {x['code'] for x in D['inbound']}
dup = sorted({x['code'] for x in buy} & cont_codes)

D['local_fees'] = local_fees
D['new_inbound'] = new_in
D['new_fees'] = new_fee
D['buy'] = buy
D['use'] = use
D['lend'] = lend
D['dup_codes'] = dup
D['mines'] = ['一矿区', '二矿区', '三矿区', '六矿区', '选厂', '维修部', '沙金']
json.dump(D, open(DJ, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

print('\n达市到港费用 %d 条，合计 %s 先令（≈ %s 元 @370）'
      % (len(local_fees), f'{sum(x["tzs"] for x in local_fees):,.0f}',
         f'{sum(x["tzs"] for x in local_fees)/370:,.0f}'))
print('  单位填错已按先令处理的:', fixed)
print('新货柜 MSKU9402161：物料 %d 行 %s 元，费用 %d 条 %s 元'
      % (len(new_in), f'{sum(x["amount"] for x in new_in):,.2f}',
         len(new_fee), f'{sum(x["cny"] for x in new_fee):,.2f}'))
print('物资流水帐：采购 %d 行 / 使用 %d 行 / 借调 %d 行，没填日期 %d 行'
      % (len(buy), len(use), len(lend), nodate))
print('物料档案：原 795 + 新增 %d = %d 个' % (n_new, len(D['items'])))
print('重叠嫌疑（货柜和本地采购都有的物料）%d 个' % len(dup))
print('矿区:', D['mines'])
