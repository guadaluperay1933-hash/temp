# -*- coding: utf-8 -*-
"""把 8 张手工对账表逐行还原成业务事件（v2 · 对账明细 9.13 版）

比 v1 多出来的东西：
  · 逐笔管理费率（mrate）—— 费率不再从单位档案取，改成业务流水手工列，历史行按原表逐笔写死
  · 逐笔税费（tax_v/s/y/i）—— 税费也不再按参数预提，改成按原表逐笔录入
  · 项目简称（从摘要「…：简称」里取）
  · 项目首现顺序（按原表 8 张表的先后 + 表内行号）
  · 每一行的原始值留档（raw），供【对账差异说明】逐笔比对
跑法：python3 parse_gk2.py
"""
import openpyxl, re, json, os, collections, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, '..', '参考', '原对账明细_9.13.xlsx')
OUT  = os.path.join(HERE, 'ev_913.json')

# 每张表的列映射：sheet -> {字段: 列号}（按原表 9.13 版逐列核对过）
M = collections.OrderedDict([
 ('德誉嘉 ', dict(unit='德誉嘉', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, recv_up=13, mfee_set=16, recv_us=17, rebate=19, rebate_got=20,
                partner_due=22, partner_paid=23)),
 ('迅驰',   dict(unit='迅驰', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, tax_v=12, tax_s=13, tax_y=14, tax_i=15, tax_due=16, tax_paid=17,
                recv_up=20, mfee_set=23, bond=24, recv_us=25, pass_due=27, pass_paid=28)),
 ('华城',   dict(unit='华城', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, tax_v=12, tax_s=13, tax_y=14, tax_i=15, tax_due=16, tax_paid=17,
                recv_up=20, mfee_set=23, recv_us=24)),
 ('金沁',   dict(unit='金沁', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, tax_v=12, tax_s=13, tax_y=14, tax_i=15, tax_due=16, tax_paid=17,
                bond_up=20, recv_up=21, bond=24, mfee_set=25, recv_us=26)),
 ('湖南锦泰', dict(unit='湖南锦泰', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, wage=11, tax_v=13, tax_s=14, tax_y=15, tax_i=16, tax_due=17,
                tax_paid=18, recv_up=21, mfee_set=24, bond=25, recv_us=26)),
 ('康欣',   dict(unit='康欣', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, tax_v=12, tax_s=13, tax_y=14, tax_i=15, tax_due=16, tax_paid=17,
                recv_up=20, bond=23, mfee_set=24, recv_us=25)),
 ('安锐',   dict(unit='安锐', proj=2, memo=3, biz=5, inv=4, to=6, sale=7, mfee=8, cost_due=9,
                cost_done=10, wage=11, tax_v=13, tax_s=14, tax_y=15, tax_i=16, tax_due=17,
                tax_paid=18, recv_up=21, mfee_set=24, bond=25, recv_us=26)),
 ('杰华电气', dict(unit='杰华', proj=2, memo=3, biz=4, inv=5, to=6, sale=7, mfee=9, cost_due=10,
                cost_done=11, wage=12, tax_v=14, tax_s=15, tax_y=16, tax_i=17, tax_due=18,
                tax_paid=19, recv_up=22, mfee_set=25, bond=26, recv_us=27)),
])
OWNERS = {'民能', '铜梁供电'}
US = {'泓普', '仟茂'}
NORM = {'德誉佳': '德誉嘉', '德誉嘉': '德誉嘉', '金沁': '金沁', '康欣': '康欣', '迅驰': '迅驰',
        '华城': '华城', '安锐': '安锐', '湖南锦泰': '湖南锦泰', '杰华': '杰华',
        '民能': '民能', '铜梁供电': '铜梁供电', '泓普': '泓普', '仟茂': '仟茂'}

def num(ws, r, c):
    if not c: return 0.0
    v = ws.cell(row=r, column=c).value
    return float(v) if isinstance(v, (int, float)) else 0.0
def txt(ws, r, c):
    if not c: return ''
    v = ws.cell(row=r, column=c).value
    return str(v).strip() if v is not None else ''

ALLU = ['泓普','仟茂','民能','铜梁供电','德誉佳','德誉嘉','迅驰','华城','康欣','金沁','安锐','湖南锦泰','杰华']
def parse_pair(memo):
    """从摘要里抓「X开…到Y」的开票方与收票方；原表收票单位列填错时以摘要为准"""
    if not memo: return None, None
    hits = []
    for u in ALLU:
        i = memo.find(u)
        while i >= 0:
            hits.append((i, NORM.get(u, u))); i = memo.find(u, i + 1)
    hits.sort()
    seen = []
    for _, u in hits:
        if not seen or seen[-1] != u: seen.append(u)
    if len(seen) >= 2: return seen[0], seen[1]
    if len(seen) == 1: return seen[0], None
    return None, None

def norm_unit(s):
    s = (s or '').strip()
    for k, v in NORM.items():
        if k in s: return v
    return s

DATE_RE = re.compile(r'(20\d{2})[.\-年/]\s*(\d{1,2})[.\-月/]\s*(\d{1,2})')
def grab_date(memo, fallback):
    m = DATE_RE.search(memo or '')
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try: return dt.date(y, mo, d)
        except ValueError: pass
    m2 = re.search(r'(20\d{2})[.\-年/]\s*(\d{1,2})月', memo or '')
    if m2:
        try: return dt.date(int(m2.group(1)), int(m2.group(2)), 1)
        except ValueError: pass
    return fallback

# ---------------- 项目全称规范化 & 简称 ----------------
def norm_proj(s):
    s = re.sub(r'\s+', '', str(s))
    s = s.replace('（', '(').replace('）', ')').replace('，', ',')
    s = re.sub(r'^项目名称[:：]', '', s)
    return s

def short_of(memo):
    """摘要形如「2026.3.27德誉佳开票到民能：平双线」，冒号后面那段就是简称"""
    m = re.sub(r'\s+', '', str(memo or ''))
    if '：' not in m and ':' not in m: return ''
    cand = re.split(r'[:：]', m)[-1].strip()
    cand = re.sub(r'[（(](土建|安装|劳务|设备|机械)[）)]', '', cand)
    cand = re.split(r'[，,]', cand)[0].strip()
    cand = re.sub(r'(台区$|台$)', '', cand)
    if re.fullmatch(r'[\d.．]+', cand): return ''
    return cand[:24]

wb = openpyxl.load_workbook(SRC, data_only=True)
events, projects = [], {}
proj_order, proj_first, short_votes = [], {}, collections.defaultdict(collections.Counter)
raw_rows = []

# ---------------------------------------------------------------- 预扫
# 原表有 84 行「项目全称」是空的。光靠沿用上一行，会把它们全挂到上面那个项目头上 ——
# 德誉嘉 11~16 行那六张成本票就是这么全堆到「平凉线惠丰9社」上的，
# 结果平凉线的已收成本票虚高几十万，平双线等五个项目一分都收不到，
# 「还差成本票」于是恰好等于管理费，看着像口径错，其实是项目挂错了。
# 摘要冒号后面本来就点名了真项目（「华城开票到德誉佳：平双线」），
# 所以先扫一遍建「简称 → 项目全称」对照表，空行照摘要归位。
SHORT2PROJ_SHEET = collections.defaultdict(dict)   # 同一个简称在各表可能指不同项目，按表优先
SHORT2PROJ_ALL   = {}
for _sh, _mp in M.items():
    _ws = wb[_sh]
    for _r in range(5, _ws.max_row + 1):
        _praw = txt(_ws, _r, _mp['proj'])
        if not _praw: continue
        _p = norm_proj(_praw)
        _s = short_of(txt(_ws, _r, _mp['memo']))
        if not _s: continue
        SHORT2PROJ_SHEET[_sh].setdefault(_s, _p)
        SHORT2PROJ_ALL.setdefault(_s, _p)
_recovered = _fellback = 0
_recover_log = []

for sh, mp in M.items():
    ws = wb[sh]; unit = mp['unit']; cur_proj = ''; fb = dt.date(2026, 1, 1)
    for r in range(5, ws.max_row + 1):
        proj_raw = txt(ws, r, mp['proj'])
        memo = txt(ws, r, mp['memo'])
        if not proj_raw and memo:
            # 项目那一格空着：先照摘要认领，认不出来才沿用上一行
            _s = short_of(memo)
            _hit = SHORT2PROJ_SHEET[sh].get(_s) or SHORT2PROJ_ALL.get(_s)
            if _hit:
                if _hit != cur_proj:
                    _recovered += 1
                    _recover_log.append((sh, r, _s, cur_proj[:18], _hit[:18]))
                cur_proj = _hit
            else:
                _fellback += 1
        if proj_raw:
            cur_proj = norm_proj(proj_raw)
            projects.setdefault(cur_proj, set()).add(unit)
            if cur_proj not in proj_first:
                proj_first[cur_proj] = (sh, r); proj_order.append(cur_proj)
            s = short_of(memo)
            if s: short_votes[cur_proj][s] += 1
        if not memo and not proj_raw: continue
        d = grab_date(memo, fb); fb = d
        to = norm_unit(txt(ws, r, mp.get('to')))
        inv = txt(ws, r, mp.get('inv'))
        biz = txt(ws, r, mp.get('biz'))
        base = dict(date=d.isoformat(), unit=unit, proj=cur_proj, memo=memo, inv=inv, biz=biz,
                    src=f'{sh}!{r}')
        sale = num(ws, r, mp['sale']); cost = num(ws, r, mp['cost_done'])
        cdue_r = round(num(ws, r, mp['cost_due']), 2)
        mfee = round(num(ws, r, mp['mfee']), 2)
        tv = round(num(ws, r, mp.get('tax_v')), 2); ts = round(num(ws, r, mp.get('tax_s')), 2)
        ty = round(num(ws, r, mp.get('tax_y')), 2); ti = round(num(ws, r, mp.get('tax_i')), 2)
        tpaid = round(num(ws, r, mp.get('tax_paid')), 2)
        # 原表「应扣税费」那一列才是权威数：四项分别取两位后的尾差补回增值税那一项，
        # 免得一分一分累出来跟原表合计对不上（迅驰就差了 3 分，看起来像「还欠税」）
        tdue_raw = round(num(ws, r, mp.get('tax_due')), 2)
        if mp.get('tax_due') and abs(tdue_raw) > 0.004:
            gap = round(tdue_raw - (tv + ts + ty + ti), 2)
            if abs(gap) > 0.0001: tv = round(tv + gap, 2)
        raw_rows.append(dict(sheet=sh, row=r, unit=unit, proj=cur_proj, memo=memo, inv=inv, biz=biz,
                             sale=round(sale, 2), mfee=mfee,
                             cost_due=round(num(ws, r, mp['cost_due']), 2),
                             cost_done=round(cost, 2),
                             tax_v=tv, tax_s=ts, tax_y=ty, tax_i=ti,
                             tax_due=round(num(ws, r, mp.get('tax_due')), 2), tax_paid=tpaid,
                             recv_up=round(num(ws, r, mp.get('recv_up')), 2),
                             recv_us=round(num(ws, r, mp.get('recv_us')), 2),
                             mfee_set=round(num(ws, r, mp.get('mfee_set')), 2),
                             bond=round(num(ws, r, mp.get('bond')), 2),
                             rebate=round(num(ws, r, mp.get('rebate')), 2),
                             wage=round(num(ws, r, mp.get('wage')), 2)))
        m_iss, m_rcv = parse_pair(memo)
        if abs(sale) > 0.004:
            si = m_iss if (m_iss and m_rcv) else unit
            sr = m_rcv if (m_iss and m_rcv) else (to or '民能')
            events.append({**base, 'kind': '销项开票', 'payer': si, 'payee': sr,
                           'amt': round(sale, 2), 'mfee': mfee,
                           'mrate': round(mfee / sale, 6) if sale else 0.0,
                           'cost_due': round(num(ws, r, mp['cost_due']), 2),
                           'tax_v': tv, 'tax_s': ts, 'tax_y': ty, 'tax_i': ti,
                           'rebate': round(num(ws, r, mp.get('rebate')), 2),
                           'nofee': 1 if abs(mfee) < 0.005 else 0})
        # 同一行两列填的是同一张票时不要记两遍
        if abs(cost) > 0.004 and abs(cost - sale) > 0.004:
            iss = m_iss or '泓普'
            rcv = m_rcv or to or unit
            # 摘要里只出现一家（「康欣代发社保…打款到康欣」「金沁罗会计对账…」这种），
            # 抓出来的开票方和收票方会是同一家，记成「康欣开票给康欣」既看不懂也进不了往来。
            # 这一列本来就是「已提供成本票」，票是我方开给挂靠单位的，直接按原意归位。
            if iss == rcv:
                iss, rcv = '泓普', unit
            # 本行自己有销项，而且「已提供成本票」正好等于本行「应收成本票」：
            # 这一格是我方补给挂靠单位的那张成本票，不是挂靠单位又往外开了一张
            # （康欣!50 的 66,500 就是这样，原来被当成第二笔销项，跟金沁表里的同一张票撞了）
            elif abs(sale) > 0.004 and abs(cost - cdue_r) <= 0.004 and abs(cdue_r) > 0.004:
                iss = next((u for u in ('仟茂', '泓普') if u in (memo or '')), '泓普')
                rcv = unit
            kind = '成本票' if iss in US else '销项开票'
            ev = {**base, 'kind': kind, 'payer': iss, 'payee': rcv, 'amt': round(cost, 2),
                  'mfee': 0, 'cost_due': 0, 'from_cost_col': True}
            if kind == '销项开票':
                ev.update(mrate=0.0, nofee=1, tax_v=0, tax_s=0, tax_y=0, tax_i=0)
            events.append(ev)
        if abs(tpaid) > 0.004:
            events.append({**base, 'kind': '已交税', 'payer': unit, 'payee': '税局',
                           'amt': tpaid, 'mfee': 0, 'cost_due': 0})
        if abs(tv + ts + ty + ti) > 0.004:
            events.append({**base, 'kind': '预提税费', 'payer': unit, 'payee': '税局',
                           'amt': round(tv + ts + ty + ti, 2), 'mfee': 0, 'cost_due': 0,
                           'tax_v': tv, 'tax_s': ts, 'tax_y': ty, 'tax_i': ti})
        ru = num(ws, r, mp.get('recv_up'))
        if abs(ru) > 0.004:
            events.append({**base, 'kind': '挂靠单位代收', 'payer': '民能', 'payee': unit,
                           'amt': round(ru, 2), 'mfee': 0, 'cost_due': 0})
        rs = num(ws, r, mp.get('recv_us'))
        if abs(rs) > 0.004:
            events.append({**base, 'kind': '我方收款', 'payer': unit, 'payee': '泓普',
                           'amt': round(rs, 2), 'mfee': 0, 'cost_due': 0})
        ms = num(ws, r, mp.get('mfee_set'))
        if abs(ms) > 0.004:
            events.append({**base, 'kind': '管理费结算', 'payer': '泓普', 'payee': unit,
                           'amt': round(ms, 2), 'mfee': 0, 'cost_due': 0})
        bd = num(ws, r, mp.get('bond'))
        if abs(bd) > 0.004:
            events.append({**base, 'kind': '扣质保金', 'payer': unit, 'payee': '泓普',
                           'amt': round(bd, 2), 'mfee': 0, 'cost_due': 0})
        # 德誉嘉表最右边那三列「合伙项目应付款」：合伙方那一份该转给人家的钱，
        # 原来一笔都没进系统，往来台账上看不到
        pd_ = num(ws, r, mp.get('partner_due'))
        if abs(pd_) > 0.004:
            events.append({**base, 'kind': '其他应付发生', 'payer': unit, 'payee': unit,
                           'amt': round(pd_, 2), 'mfee': 0, 'cost_due': 0,
                           'memo': (memo or '合伙项目应付款') + '（合伙项目应付款）'})
        pp_ = num(ws, r, mp.get('partner_paid'))
        if abs(pp_) > 0.004:
            events.append({**base, 'kind': '其他应付支付', 'payer': unit, 'payee': unit,
                           'amt': round(pp_, 2), 'mfee': 0, 'cost_due': 0,
                           'memo': (memo or '合伙项目应付款') + '（已付合伙方）'})

# ---------------- 逐表把税费尾差配平到原表合计行 ----------------
# 原表每一行的税费小数位比显示的多，逐行取两位再相加会跟合计行差上一两分，
# 看起来就像「明明交清了还欠两分税」。这里按每张表把尾差补到该表最大的那一笔上。
for sh, mp in M.items():
    ws = wb[sh]
    for col_key, ev_kind in (('tax_due', None), ('tax_paid', '已交税')):
        c = mp.get(col_key)
        if not c: continue
        tgt = ws.cell(row=4, column=c).value
        if not isinstance(tgt, (int, float)): continue
        tgt = round(float(tgt), 2)
        if ev_kind:
            rows = [e for e in events if e['kind'] == ev_kind and e['src'].split('!')[0] == sh]
            cur = round(sum(e['amt'] for e in rows), 2)
            if rows and abs(cur - tgt) > 0.004:
                big = max(rows, key=lambda e: abs(e['amt']))
                big['amt'] = round(big['amt'] + (tgt - cur), 2)
        else:
            rows = [e for e in events if e['kind'] == '预提税费' and e['src'].split('!')[0] == sh]
            cur = round(sum(e['tax_v'] + e['tax_s'] + e['tax_y'] + e['tax_i'] for e in rows), 2)
            if rows and abs(cur - tgt) > 0.004:
                big = max(rows, key=lambda e: abs(e['tax_v']))
                big['tax_v'] = round(big['tax_v'] + (tgt - cur), 2)
    # 销项开票行上挂的那份税费跟着「预提税费」事件走，逐行同步回去
    fix = {e['src']: e for e in events if e['kind'] == '预提税费' and e['src'].split('!')[0] == sh}
    for e in events:
        if e['kind'] == '销项开票' and not e.get('from_cost_col') and e['src'] in fix:
            f = fix[e['src']]
            e['tax_v'], e['tax_s'] = f['tax_v'], f['tax_s']
            e['tax_y'], e['tax_i'] = f['tax_y'], f['tax_i']

# 跨表重复：同一张票在两张表里各记一次，只留信息最全的那一条计入汇总。
# 「信息最全」＝不是从成本票列还原出来的 > 记了管理费的 > 记了税费的 > 记了应到成本票的 > 先出现的。
def _info(e):
    return (0 if e.get('from_cost_col') else 1,
            1 if abs(e.get('mfee') or 0) > 0.004 else 0,
            1 if abs(sum(float(e.get(k) or 0) for k in ('tax_v', 'tax_s', 'tax_y', 'tax_i'))) > 0.004 else 0,
            1 if abs(float(e.get('cost_due') or 0)) > 0.004 else 0)
groups = {}
for i, e in enumerate(events):
    if e['kind'] not in ('成本票', '销项开票'):
        e['count_in'] = '是'; continue
    k = (e['kind'], e['payer'], e['payee'], round(e['amt'], 2))
    groups.setdefault(k, []).append(i)
for k, idxs in groups.items():
    by_sheet = {}
    for i in idxs:
        by_sheet.setdefault(events[i]['src'].split('!')[0], []).append(i)
    if len(by_sheet) == 1:
        for i in idxs: events[i]['count_in'] = '是'      # 同一张表内同额多笔＝确实是不同的票
        continue
    # 每张表各留一条候选，再在候选里挑信息最全的那条当主记录
    cands = [v[0] for v in by_sheet.values()]
    best = max(cands, key=lambda i: (_info(events[i]), -i))
    for i in idxs:
        if i == best or events[i]['src'].split('!')[0] == events[best]['src'].split('!')[0]:
            events[i]['count_in'] = '是'
        else:
            events[i]['count_in'] = '否'; events[i]['dup_of'] = events[best]['src']
    # 同一张表里多出来的那几条也要标否（同表同额多笔本来算不同的票，但跨表重复组里只保主表）
    for sheet, v in by_sheet.items():
        if sheet == events[best]['src'].split('!')[0]:
            for i in v: events[i]['count_in'] = '是'

# ---------------- 项目简称：一个全称一个简称，重名的后面补一级单位 ----------------
# 摘要里同一个项目有好几种写法时，取「短的、口语的」那一个
SHORT_FIX = {
    '侣新线铜安1社台': '侣新线铜安1社', '平镇线新桥社台': '平镇线新桥社',
    '侣新线凤飞7': '侣新线凤飞7社', '黄荆4社和三湾1.2社': '黄荆4社',
    '少盘线哨楼村支线#1杆': '少盘线哨楼村支线', '少盘线哨楼村支线#1杆线路迁改': '少盘线哨楼村支线',
    '永嘉所义和8社等台区': '永嘉所义和8社', '永嘉所许家3社1等台区': '永嘉所许家3社',
    '永嘉所许家3社1等': '永嘉所许家3社', '永嘉所高丰3社等台区': '永嘉所高丰3社',
    '永嘉所高丰3社等': '永嘉所高丰3社', '大太线老旧线路': '大太线',
    '大太线老旧线路代发工资17490': '大太线', '金湖东西': '金湖东西线',
    '金湖东西电缆改造': '金湖东西线', '金湖东西（安装）': '金湖东西线',
    '永嘉所10kV西新线等2条线路': '永嘉所10kV西新线',
    '输电运检中心2026年': '输电运检中心线路接地隐患消除',
    '华润电力平滩农光互补项目220kV升压站PC工程': '平滩农光互补219kV',
    '华润电力滩农光互补项目220kV升压站PC工程': '平滩农光互补220kV',
    '电力新建（康欣3W': '园区光伏电力新建', '电力新建工程（康欣3W': '园区光伏电力新建',
    '桥亭水云居（项目8W': '桥亭水云居', '特种设备租赁服务': '峰盛模具箱变拆除',
    '锦绣花园A区': '锦绣花园A区', '永嘉所10kV西新线': '永嘉所10kV西新线',
}
raw_short = {}
for p in proj_order:
    c = short_votes.get(p)
    s0 = c.most_common(1)[0][0] if c else ''
    if not s0:
        s0 = re.sub(r'^重庆(市)?铜梁?(区)?', '', p)[:12]
    raw_short[p] = SHORT_FIX.get(s0, s0)
# 重名的补「(一级单位)」，还重就补 -1 -2
cnt0 = collections.Counter(raw_short.values())
shorts, used = {}, collections.Counter()
for p in proj_order:
    s = raw_short[p]
    if cnt0[s] > 1:
        s = f"{s}({M[proj_first[p][0]]['unit']})"      # 用「原表里第一次出现在哪张表」来消歧
    while s in used:
        used[s] += 1
        s = f'{raw_short[p]}-{used[raw_short[p]] + 1}'
    used[s] += 1
    shorts[p] = s
dupe_short = {s for s, n in collections.Counter(shorts.values()).items() if n > 1}

data = dict(events=events,
            proj_order=[dict(full=p, short=shorts[p], first_sheet=proj_first[p][0],
                             first_row=proj_first[p][1], units=sorted(projects.get(p, [])),
                             short_dupe=1 if shorts[p] in dupe_short else 0)
                        for p in proj_order],
            raw_rows=raw_rows)
json.dump(data, open(OUT, 'w'), ensure_ascii=False)

print(f'项目归位：照摘要认回 {_recovered} 行，认不出来沿用上一行 {_fellback} 行')
for _x in _recover_log[:8]:
    print('   %s!%d 「%s」 %s → %s' % _x)
print('事件总数', len(events))
print('按类型:', dict(collections.Counter(e['kind'] for e in events)))
print('跨表重复标记', sum(1 for e in events if e.get('count_in') == '否'), '笔')
print('项目数', len(proj_order), '｜简称重名', len(dupe_short), '个：', '、'.join(sorted(dupe_short)))
print('原表留档行', len(raw_rows))

ORIG = {'德誉嘉': (651402.95, 611391.37), '迅驰': (219510.99, 210730.55),
        '华城': (472843.99, 463387.12), '康欣': (1020420.18, 943853.98),
        '金沁': (600622.61, 581152.54), '安锐': (141749.77, 133690.89),
        '湖南锦泰': (163893.26, 0.0), '杰华': (122305.72, 110075.15)}
V = [e for e in events if e.get('count_in') == '是']
print('\n各单位核对（已排除跨表重复）：')
for u, (os_, oc) in ORIG.items():
    sale = sum(e['amt'] for e in V if e['kind'] == '销项开票' and e['payer'] == u)
    recv = sum(e['amt'] for e in V if e['payee'] == u and e['kind'] in ('成本票', '销项开票'))
    print(f'  {u:<5} 开票 {sale:>12,.2f} vs {os_:>12,.2f} {"OK " if abs(sale-os_)<1 else "差异"}'
          f' | 收到成本票 {recv:>12,.2f} vs {oc:>12,.2f} {"OK " if abs(recv-oc)<1 else "差异"}')
print('\n税费核对（原表逐行 P/Q 列合计）：')
TAXO = {'迅驰': (10595.25, 10595.25), '华城': (27012.71, 16883.46), '金沁': (28854.13, 11075.59),
        '湖南锦泰': (15156.37, 17766.62), '康欣': (35030.80, 18726.14), '安锐': (332.72, 332.72),
        '杰华': (15759.04, 0.0)}
for u, (td, tp) in TAXO.items():
    mine_d = sum(e.get('tax_v', 0) + e.get('tax_s', 0) + e.get('tax_y', 0) + e.get('tax_i', 0)
                 for e in events if e['kind'] == '预提税费' and e['unit'] == u)
    mine_p = sum(e['amt'] for e in events if e['kind'] == '已交税' and e['unit'] == u)
    print(f'  {u:<5} 应扣税费 {mine_d:>11,.2f} vs {td:>11,.2f} {"OK " if abs(mine_d-td)<1 else "差异"}'
          f' | 已交税 {mine_p:>11,.2f} vs {tp:>11,.2f} {"OK " if abs(mine_p-tp)<1 else "差异"}')
