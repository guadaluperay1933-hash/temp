# -*- coding: utf-8 -*-
"""把《集装箱库存明细表》877 行清洗成物料档案 + 入库单 + 费用台账。

几条要紧的清洗规则（都是拿真实数据验过的）：
1. 集装箱号和供货商列有 394 行写「同上」，必须向下填充还原，否则一排序整批货就乱了。
2. 费用行（海运费/服务费/装柜费/叉车费/超时费/餐费/运费）混在物料行中间，
   靠 A 列或供货商认不全 —— r724 海运费 22835、r725 服务费 3000 的 A 列是空的，
   供货商写的是「福建全力供应链」，和货物同一家。所以按**品名关键字**认，单独拉出来。
3. 单位里有 342 行是「数字+文字」（1台/10个/100米…）。其中 335 行那个数字正好等于入库数量，
   是把数量重复写进了单位，去掉数字即可；另外 7 行不等（100米/2米/支），
   那个数字是**包装规格**，必须原样保留 —— 这时单价是「每包装」的价。
   所以本系统不做单位换算：数量、单位、单价三者一致，金额 = 数量 × 单价，跟原表 M 列口径一致。
4. 物料按（品名+规格）去重发编码，跨表只用编码关联，不用品名 ——
   「钢球 Φ30」和「钢球-30」是同一个东西两种写法，用名字关联必错。
"""
import json, os, re, sys, datetime, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, '参考', '原_集装箱库存明细表.xlsx')
OUT = os.path.join(HERE, 'mine_data.json')
SHEET = '按照这个为准带价格'

FEE_WORDS = ('海运费', '服务费', '搬运费', '叉车费', '超时费', '餐费', '装柜', '报关', '清关',
             '港杂', '拖车费', '吊柜')
FEE_CATS = {'装柜费用', '运费', '海运费'}

CLASS = [
    (('炸药', '雷管', '导爆', '乳化', '硝铵', '起爆', '导火'), '炸药火工'),
    (('柴油', '汽油', '机油', '液压油', '齿轮油', '黄油', '润滑', '防冻液'), '油料'),
    (('黄原酸', '硫酸铜', '碳酸钠', '松醇油', '硅酸钠', '硫化钠', 'JB-05', '二丁基',
      '药剂', '起泡剂', '捕收剂', '水玻璃', '石灰'), '选矿药剂'),
    (('钢球', '钢段', '钢锻'), '磨矿介质'),
    (('钢板', '花纹板', '焊管', '无缝管', '角钢', '扁钢', '钢轨', '圆钢', '槽钢',
      '工字钢', '方管', '钢管', '镀锌', '钢材', '型钢'), '钢材'),
    (('电缆', '电线', '屏蔽线', '开关', '启动柜', '变压器', '配电', '电柜', '控制柜',
      '接触器', '断路器', '电机', '灯'), '电缆电气'),
    (('球磨机', '破碎机', '浮选机', '分级机', '绞车', '空压机', '凿岩机', '水泵', '风机',
      '通风机', '叉车', '地磅', '发电机', '皮带秤', '压滤机', '给矿机', '振动筛',
      '搅拌', '装载机', '汽车衡'), '机械设备'),
    (('衬板', '叶轮', '轴承', '皮带', '滤芯', '密封', '刮板', '齿圈', '联轴器', '稳钉',
      '护板', '破碎壁', '轧臼壁', '扎臼壁', '刹车带', '配件', '备件'), '机械配件'),
    (('螺栓', '螺母', '螺丝', '垫片', '金属垫', '胶垫', '球阀', '闸阀', '法兰', '对丝',
      '弯头', '三通', '软连接', '阀', '管件', '扳手', '钳', '锤', '锯', '钻头'), '五金管件'),
    (('手套', '工作服', '安全帽', '雨鞋', '口罩', '劳保', '棉被', '毛巾', '雨衣'), '劳保用品'),
    (('集装箱',), '集装箱'),
    (('土工膜', '吨袋', '编织袋', '篷布'), '包装材料'),
]
KEY_CATS = {'炸药火工', '油料', '选矿药剂', '磨矿介质'}


def classify(name, spec, cat_hint):
    s = f'{name} {spec} {cat_hint}'
    for keys, c in CLASS:
        for k in keys:
            if k in s:
                return c
    return '其他'


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.date):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime('%Y-%m-%d')
    return ''


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def txt(v):
    return ('' if v is None else str(v)).strip()


def clean_unit(unit, qty):
    """单位里混进来的数字：等于数量的是重复写，删掉；不等于的是包装规格，留着。"""
    u = unit.strip()
    m = re.match(r'^(\d+(?:\.\d+)?)\s*(\D.*)$', u)
    if not m:
        return u, False
    n, rest = float(m.group(1)), m.group(2).strip()
    if abs(n - qty) < 1e-9 and rest:
        return rest, True        # 「10个」+数量10 → 个
    return u, False              # 「100米」+数量1 → 保留 100米（单价是每 100 米的价）


wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb[SHEET]

raw = []
last_cont = last_supp = ''
for r in range(9, 1067):
    name = txt(ws.cell(r, 6).value)
    if not name:
        continue
    cont = txt(ws.cell(r, 5).value)
    supp = txt(ws.cell(r, 4).value)
    if cont in ('同上', '"', '〃'):
        cont = last_cont
    elif cont:
        last_cont = cont
    if supp in ('同上', '"', '〃'):
        supp = last_supp
    elif supp:
        last_supp = supp
    raw.append(dict(
        row=r, date=as_date(ws.cell(r, 3).value), supplier=supp, container=cont,
        name=name, spec=txt(ws.cell(r, 7).value), unit=txt(ws.cell(r, 8).value),
        batchno=txt(ws.cell(r, 9).value), qty=num(ws.cell(r, 11).value),
        price=num(ws.cell(r, 12).value), amount=num(ws.cell(r, 13).value),
        qout=num(ws.cell(r, 14).value), note=txt(ws.cell(r, 21).value),
        hint=txt(ws.cell(r, 1).value)))

fees, goods, outs = [], [], []
for x in raw:
    is_fee = x['hint'] in FEE_CATS or any(w in x['name'] for w in FEE_WORDS)
    if is_fee:
        fees.append(x)
    elif x['qty'] == 0 and x['qout'] > 0:
        # 原表第 439 行：空集装箱卖给三矿。全表唯一一条真出库记录，
        # 入库数量是 0、出库数量是 1，不能当入库行处理。
        outs.append(x)
    else:
        goods.append(x)

# ── 单位清洗 ──────────────────────────────────────────────────────────
n_strip = n_keep = 0
for x in goods:
    u, stripped = clean_unit(x['unit'], x['qty'])
    x['unit_clean'] = u or '个'
    if stripped:
        n_strip += 1
    elif re.match(r'^\d', x['unit']):
        n_keep += 1
        x['note'] = (x['note'] + ' 单位含包装规格，单价是每「%s」的价' % x['unit']).strip()

# ── 物料去重发编码 ────────────────────────────────────────────────────
items, key2code = [], {}
for x in goods:
    key = (x['name'], x['spec'])
    if key not in key2code:
        code = 'JH%04d' % (len(items) + 1)
        key2code[key] = code
        items.append(dict(code=code, name=x['name'], spec=x['spec'],
                          unit=x['unit_clean'],
                          cat=classify(x['name'], x['spec'], x['hint']),
                          key='是' if classify(x['name'], x['spec'], x['hint']) in KEY_CATS else '否'))
    x['code'] = key2code[key]

# ── 批次：柜号 + 到港年月 做唯一键（柜号会被船公司重复使用）────────────
for x in goods + fees:
    c = x['container'] or '本地采购'
    ym = x['date'][:7].replace('-', '') if x['date'] else '000000'
    x['batch'] = f'{c}-{ym}' if x['container'] else f'本地-{ym}'
    x['source'] = 'A国内' if x['container'] else 'B当地'

batches = collections.OrderedDict()
for x in goods:
    b = batches.setdefault(x['batch'], dict(batch=x['batch'], container=x['container'],
                                            date=x['date'], source=x['source'],
                                            rows=0, value=0.0, nopricerows=0))
    b['rows'] += 1
    b['value'] += x['amount']
    if x['price'] == 0:
        b['nopricerows'] += 1
for x in fees:
    if x['batch'] in batches:
        batches[x['batch']]['fee'] = batches[x['batch']].get('fee', 0.0) + x['amount']

for x in outs:
    key = (x['name'], x['spec'])
    if key not in key2code:
        code = 'JH%04d' % (len(items) + 1)
        key2code[key] = code
        items.append(dict(code=code, name=x['name'], spec=x['spec'], unit=x['unit'] or '个',
                          cat=classify(x['name'], x['spec'], x['hint']), key='否'))
    x['code'] = key2code[key]
    x['unit_clean'] = x['unit'] or '个'

data = dict(items=items, inbound=goods, outbound=outs, fees=fees, batches=list(batches.values()))
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

print('原始行 %d → 入库物料行 %d，出库行 %d，费用行 %d' % (len(raw), len(goods), len(outs), len(fees)))
print('去重后物料 %d 个' % len(items))
print('单位清洗：删掉重复数字 %d 行，保留包装规格 %d 行' % (n_strip, n_keep))
print('物料货值合计 %.2f 元，费用合计 %.2f 元' % (sum(x['amount'] for x in goods),
                                                 sum(x['amount'] for x in fees)))
print('没填单价的物料行 %d / %d' % (sum(1 for x in goods if x['price'] == 0), len(goods)))
print('批次 %d 个，其中有费用的 %d 个' % (len(batches), sum(1 for b in batches.values() if b.get('fee'))))
print('\n大类分布:', dict(collections.Counter(i['cat'] for i in items).most_common()))
print('\n有费用的批次:')
for b in batches.values():
    if b.get('fee'):
        rate = b['fee'] / b['value'] if b['value'] else 0
        print('   %-28s 货值 %12.2f  费用 %10.2f  费率 %6.2f%%  无价行 %d/%d'
              % (b['batch'], b['value'], b['fee'], rate * 100, b['nopricerows'], b['rows']))
