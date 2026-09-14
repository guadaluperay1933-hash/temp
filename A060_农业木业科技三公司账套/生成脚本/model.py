# -*- coding: utf-8 -*-
"""A060 · 纯 Python 复算模型（内账含税口径）

跟 Excel 公式完全独立地把同一批业务再算一遍：存货加权平均、批次成本分摊、
借贷分录、科目余额、资产负债表、利润表。verify 用它逐项比对表里算出来的结果。

口径：收入/成本/费用一律含税，不拆进项销项；税按【资金流水】里「缴纳税费」
实际缴的钱，当期计入费用科目。
"""
from collections import defaultdict
import a060_data as D

R2 = lambda x: round(x + (1e-9 if x >= 0 else -1e-9), 2)

GD = {g[0]: dict(code=g[0], name=g[1], spec=g[2], unit=g[3], co=g[4], cls=g[5],
                 inv=g[6], rev=g[7], cost=g[8], rate=g[9]) for g in D.GOODS}
ACC = {a[1]: dict(code=a[0], name=a[1], cls=a[2], dir=a[3], rpt=a[4], item=a[5]) for a in D.ACCS}
BANK = {a[1]: dict(code=a[0], name=a[1], kind=a[4], acc=a[5], co=a[6], open=a[7]) for a in D.ACCOUNTS}
key = lambda co, gc: (co, gc)


def build():
    m = {}
    # ---------- 购销流水（含税金额 = 数量 × 含税单价） ----------
    buy = []
    for i, (d, co, kind, pt, gc, qty, price, inv, invno, memo) in enumerate(D.EX_BUY):
        buy.append(dict(r=i, date=d, co=co, kind=kind, pt=pt, gc=gc, qty=qty, price=price,
                        inv=inv, invno=invno, memo=memo, amt=R2(qty * price), cost=0.0))
    m['buy'] = buy

    # ---------- 存货加权平均 + 批次成本 ----------
    o_qty, o_amt = defaultdict(float), defaultdict(float)
    for co, gc, q, a in D.OPEN_INV:
        o_qty[key(co, gc)] += q; o_amt[key(co, gc)] += a
    p_qty, p_amt = defaultdict(float), defaultdict(float)
    for b in buy:
        sg = 1 if b['kind'] == '采购入库' else (-1 if b['kind'] == '采购退回' else 0)
        if sg:
            p_qty[key(b['co'], b['gc'])] += sg * b['qty']
            p_amt[key(b['co'], b['gc'])] += sg * b['amt']

    prod = [dict(r=i, date=d, co=co, bt=bt, kind=kind, gc=gc, qty=qty, fee=fee,
                 ctr=ctr, ei=ei, w=w, memo=memo, amt=0.0)
            for i, (d, co, bt, kind, gc, qty, fee, ctr, ei, w, memo) in enumerate(D.EX_PROD)]
    batches = []
    for p in prod:
        if (p['co'], p['bt']) not in batches: batches.append((p['co'], p['bt']))

    m_qty, m_amt = defaultdict(float), defaultdict(float)
    avg = {}
    def recompute_avg(k):
        q = o_qty[k] + p_qty[k] + m_qty[k]
        a = o_amt[k] + p_amt[k] + m_amt[k]
        avg[k] = (a / q) if abs(q) > 1e-9 else 0.0   # 不取整，避免大数量品种累出尾差

    produced_by = defaultdict(list)
    for p in prod:
        if p['kind'] == '产成品入库':
            produced_by[key(p['co'], p['gc'])].append((p['co'], p['bt']))

    done, guard = set(), 0
    while len(done) < len(batches) and guard < 100:
        guard += 1
        for bk in batches:
            if bk in done: continue
            rows = [p for p in prod if (p['co'], p['bt']) == bk]
            ins = [p for p in rows if p['kind'] == '领用投入']
            if any(b2 not in done for p in ins for b2 in produced_by[key(p['co'], p['gc'])]):
                continue
            for p in ins:
                recompute_avg(key(p['co'], p['gc']))
                p['amt'] = R2(p['qty'] * avg[key(p['co'], p['gc'])])
            for p in rows:
                if p['kind'] == '加工费用': p['amt'] = R2(p['fee'])
            tot = R2(sum(p['amt'] for p in rows if p['kind'] != '产成品入库'))
            outs = [p for p in rows if p['kind'] == '产成品入库']
            wsum = sum((p['w'] or 1) for p in outs)
            cum = 0.0
            for p in outs:                                   # 累计权重法：最后一条吸收分位差
                w = p['w'] or 1
                cum += w
                p['amt'] = (R2(tot * cum / wsum) - R2(tot * (cum - w) / wsum)) if wsum else 0.0
                m_qty[key(p['co'], p['gc'])] += p['qty']
                m_amt[key(p['co'], p['gc'])] += p['amt']
            done.add(bk)
    assert len(done) == len(batches), '生产批次存在循环依赖'
    for k in set(list(o_qty) + list(p_qty) + list(m_qty)):
        recompute_avg(k)
    m['prod'] = prod; m['avg'] = avg
    m['inv'] = {k: dict(oq=o_qty[k], oa=o_amt[k], pq=p_qty[k], pa=p_amt[k],
                        mq=m_qty[k], ma=m_amt[k], avg=avg.get(k, 0.0))
                for k in set(list(o_qty) + list(p_qty) + list(m_qty))}

    for b in buy:
        if b['kind'] in ('销售出库', '销售退回'):
            b['cost'] = R2(b['qty'] * avg.get(key(b['co'], b['gc']), 0.0))

    s_qty, s_amt, u_qty, u_amt = (defaultdict(float) for _ in range(4))
    for b in buy:
        sg = 1 if b['kind'] == '销售出库' else (-1 if b['kind'] == '销售退回' else 0)
        if sg:
            s_qty[key(b['co'], b['gc'])] += sg * b['qty']
            s_amt[key(b['co'], b['gc'])] += sg * b['cost']
    for p in prod:
        if p['kind'] == '领用投入':
            u_qty[key(p['co'], p['gc'])] += p['qty']; u_amt[key(p['co'], p['gc'])] += p['amt']
    for k, v in m['inv'].items():
        v['sq'], v['sa'], v['uq'], v['ua'] = s_qty[k], s_amt[k], u_qty[k], u_amt[k]
        v['eq'] = v['oq'] + v['pq'] + v['mq'] - v['sq'] - v['uq']
        v['ea'] = R2(v['oa'] + v['pa'] + v['ma'] - v['sa'] - v['ua'])

    # ---------- 资金流水 ----------
    cash = [dict(r=i, date=d, co=co, acct=acct, kind=kind, pt=pt, ctr=ctr, ei=ei,
                 cin=R2(cin), cout=R2(cout), ref=ref, memo=memo, amt=R2(cin + cout))
            for i, (d, co, acct, kind, pt, ctr, ei, cin, cout, ref, memo) in enumerate(D.EX_CASH)]
    m['cash'] = cash

    # ---------- 实缴税费（按税种） ----------
    tax = defaultdict(lambda: defaultdict(float))
    for c in cash:
        if c['kind'] == '缴纳税费':
            tax[c['co']][c['ei'] or '其他税费'] += c['amt']
    m['tax'] = {co: {k: R2(v) for k, v in d_.items()} for co, d_ in tax.items()}

    oth = [dict(r=i, date=d, co=co, memo=memo, dr=dr, cr=cr, amt=amt, pt=pt, ei=ei)
           for i, (d, co, memo, dr, cr, amt, pt, ei) in enumerate(D.EX_OTH)]
    m['oth'] = oth

    # ---------- 分录 ----------
    rb = {r[0]: r[2] for r in D.RULE_BUY}
    rc = {r[0]: r[2] for r in D.RULE_CASH}
    rp = {r[0]: r[2] for r in D.RULE_PROD}
    legs = []
    for b in buy:
        g = GD[b['gc']]
        settle = '应收账款' if b['kind'].startswith('销售') else '应付账款'
        amts = {'金额': b['amt'], '成本金额': b['cost']}
        for acc, dr, code in rb[b['kind']]:
            if not acc: continue
            a = {'@结算': settle, '@存货': g['inv'], '@收入': g['rev'], '@成本': g['cost']}.get(acc, acc)
            legs.append(dict(src='购销流水', co=b['co'], date=b['date'], acc=a, pt=b['pt'],
                             dr=amts[code] if dr == '借' else 0.0,
                             cr=amts[code] if dr == '贷' else 0.0))
    for c in cash:
        for acc, dr, code in rc[c['kind']]:
            if not acc: continue
            a = {'@账户': BANK[c['acct']]['acc'], '@对方': c['ctr']}.get(acc, acc)
            legs.append(dict(src='资金流水', co=c['co'], date=c['date'], acc=a, pt=c['pt'],
                             dr=c['amt'] if dr == '借' else 0.0,
                             cr=c['amt'] if dr == '贷' else 0.0))
    for p in prod:
        g = GD.get(p['gc'], {})
        for acc, dr, code in rp[p['kind']]:
            if not acc: continue
            a = {'@存货': g.get('inv', ''), '@对方': p['ctr']}.get(acc, acc)
            legs.append(dict(src='生产加工', co=p['co'], date=p['date'], acc=a, pt='',
                             dr=p['amt'] if dr == '借' else 0.0,
                             cr=p['amt'] if dr == '贷' else 0.0))
    for o in oth:
        legs.append(dict(src='其他分录', co=o['co'], date=o['date'], acc=o['dr'], pt=o['pt'],
                         dr=o['amt'], cr=0.0))
        legs.append(dict(src='其他分录', co=o['co'], date=o['date'], acc=o['cr'], pt=o['pt'],
                         dr=0.0, cr=o['amt']))
    legs = [x for x in legs if abs(x['dr']) > 1e-9 or abs(x['cr']) > 1e-9]
    m['legs'] = legs

    # ---------- 科目余额 ----------
    tb = defaultdict(lambda: dict(od=0.0, oc=0.0, d=0.0, c=0.0))
    for co, acc, d_, c_ in D.OPEN_ACC:
        tb[(co, acc)]['od'] += d_; tb[(co, acc)]['oc'] += c_
    for x in legs:
        tb[(x['co'], x['acc'])]['d'] += x['dr']; tb[(x['co'], x['acc'])]['c'] += x['cr']
    for k, v in tb.items():
        a = ACC[k[1]]
        v['open'] = R2(v['od'] - v['oc']) if a['dir'] == '借' else R2(v['oc'] - v['od'])
        v['net'] = R2(v['d'] - v['c']) if a['dir'] == '借' else R2(v['c'] - v['d'])
        v['end'] = R2(v['open'] + v['net'])
    m['tb'] = tb

    bal = {k: v['open'] for k, v in BANK.items()}
    for c in cash: bal[c['acct']] = R2(bal[c['acct']] + c['cin'] - c['cout'])
    m['acctbal'] = bal

    # ---------- 报表 ----------
    m['pl'], m['bs'] = {}, {}
    for co in D.CO_NAMES:
        item = defaultdict(float)
        pl = defaultdict(float)
        for k, v in tb.items():
            if k[0] != co: continue
            item[ACC[k[1]]['item']] += v['end']
            if ACC[k[1]]['rpt'] == '利润表': pl[ACC[k[1]]['item']] += v['net']
        pl = {k: R2(v) for k, v in pl.items()}
        pl['营业利润'] = R2(pl.get('营业收入', 0) - pl.get('营业成本', 0) - pl.get('税金及附加', 0)
                        - pl.get('销售费用', 0) - pl.get('管理费用', 0) - pl.get('财务费用', 0)
                        - pl.get('研发费用', 0) + pl.get('投资收益', 0))
        pl['利润总额'] = R2(pl['营业利润'] + pl.get('营业外收入', 0) - pl.get('营业外支出', 0))
        pl['净利润'] = R2(pl['利润总额'] - pl.get('所得税费用', 0))
        m['pl'][co] = pl

        流动资产 = R2(item['货币资金'] + item['应收账款'] + item['预付款项'] + item['其他应收款'] + item['存货'])
        固净 = R2(item['固定资产原价'] - item['累计折旧'])
        生物净 = R2(item['生产性生物资产原价'] - item['生物资产累计折旧'])
        无形净 = R2(item['无形资产原价'] - item['累计摊销'])
        非流动 = R2(固净 + 生物净 + 无形净 + item['长期待摊费用'])
        资产 = R2(流动资产 + 非流动)
        流动负债 = R2(item['短期借款'] + item['应付账款'] + item['预收款项'] + item['应付职工薪酬']
                    + item['应交税费'] + item['其他应付款'])
        负债 = R2(流动负债 + item['长期借款'])
        未分配 = R2(item['未分配利润'] + pl['净利润'])
        权益 = R2(item['实收资本'] + item['资本公积'] + item['盈余公积'] + 未分配)
        m['bs'][co] = dict(item={k: R2(v) for k, v in item.items()}, 流动资产=流动资产,
                           固定资产净值=固净, 生物资产净值=生物净, 无形资产净值=无形净,
                           非流动资产=非流动, 资产总计=资产, 流动负债=流动负债, 负债合计=负债,
                           未分配利润=未分配, 权益合计=权益, 负债和权益=R2(负债 + 权益),
                           差额=R2(资产 - 负债 - 权益))
    return m


def period_net(m, start, end):
    """按日期区间把分录汇总成 {(公司, 科目): 本期发生净额（按科目方向取正）}"""
    t = defaultdict(lambda: [0.0, 0.0])
    for x in m['legs']:
        if start <= x['date'] <= end:
            t[(x['co'], x['acc'])][0] += x['dr']; t[(x['co'], x['acc'])][1] += x['cr']
    return {k: (R2(d_ - c_) if ACC[k[1]]['dir'] == '借' else R2(c_ - d_)) for k, (d_, c_) in t.items()}


if __name__ == '__main__':
    m = build()
    for co in D.CO_NAMES:
        p, b = m['pl'][co], m['bs'][co]
        print(f"{co}: 含税收入 {p.get('营业收入',0):>12,.2f}  含税成本 {p.get('营业成本',0):>12,.2f}  "
              f"税费 {p.get('税金及附加',0):>9,.2f}  利润总额 {p['利润总额']:>11,.2f}  "
              f"净利 {p['净利润']:>11,.2f}  资产 {b['资产总计']:>12,.2f}  平衡差 {b['差额']:.2f}")
    print('实缴税费：', {k: v for k, v in m['tax'].items()})
    print('分录条数', len(m['legs']))
    print('账户余额：', {k: round(v, 2) for k, v in m['acctbal'].items()}, '合计', round(sum(m['acctbal'].values()), 2))
    cash_by_co = {co: R2(sum(v['end'] for k, v in m['tb'].items()
                             if k[0] == co and k[1] in ('库存现金', '银行存款'))) for co in D.CO_NAMES}
    print('各公司货币资金：', cash_by_co, '合计', round(sum(cash_by_co.values()), 2))
