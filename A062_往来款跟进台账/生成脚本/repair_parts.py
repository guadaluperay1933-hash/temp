# -*- coding: utf-8 -*-
"""openpyxl 存盘会顺手扔掉几样它不认识的东西，这里原样补回去。

被扔掉的：
1. **xl/metadata.xml** 和 6 个格子上的 `cm="1"`。
   这两样是 Excel 365 用来标记「这是动态数组公式」的。现金/微信/支付宝/杭州/农商
   五张账户分表的 A4 用的是 `_xlfn._xlws.FILTER(...)`，靠 cm="1" 才知道它能往下铺开；
   标记一掉，公式还在，但退化成尺寸写死的老式 CSE 数组，只剩 2 行 —— 那五张表就废了。
2. **xl/comments1.xml / vmlDrawing1.vml**：【汇报表】I26 上一条批注。

跑法：python3 repair_parts.py <改造后的xlsx> <原始xlsx>
"""
import os, re, sys, shutil, zipfile

def _attrs(tag):
    """把一个 XML 标签里的属性抠成 dict —— 属性顺序各家写法不同，不能按顺序正则匹配。
       原件是 Id 在前 Target 在后，openpyxl 存出来正好相反，按顺序写就全都匹配不上。"""
    return dict(re.findall(r'([A-Za-z:]+)="([^"]*)"', tag))

def _norm(target):
    """Target 可能写成 worksheets/sheet1.xml，也可能写成 /xl/worksheets/sheet1.xml"""
    t = target.lstrip('/')
    return t if t.startswith('xl/') else 'xl/' + t

def sheet_map(z):
    """工作表名 -> xl/worksheets/sheetN.xml（各文件里编号不一样，只能按名字认）"""
    wb = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid2t = {}
    for m in re.finditer(r'<Relationship[^>]*/?>', rels):
        a = _attrs(m.group(0))
        if a.get('Id') and a.get('Target'): rid2t[a['Id']] = _norm(a['Target'])
    out = {}
    for m in re.finditer(r'<sheet[^>]*/?>', wb):
        a = _attrs(m.group(0))
        t = rid2t.get(a.get('r:id', ''))
        if a.get('name') and t and '/worksheets/' in t: out[a['name']] = t
    return out

def repair(dst, src):
    zs = zipfile.ZipFile(src)
    zd = zipfile.ZipFile(dst)
    smap_s, smap_d = sheet_map(zs), sheet_map(zd)
    parts = {n: zd.read(n) for n in zd.namelist()}
    zd.close()
    done = []

    # ---- ① 哪些表的哪些格子原来带 cm="1" ----
    want = {}                                   # 工作表名 -> {格子坐标}
    for nm, path in smap_s.items():
        try: x = zs.read(path).decode('utf8')
        except KeyError: continue
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*cm="1"', x):
            want.setdefault(nm, set()).add(m.group(1))

    # ---- ② metadata.xml 本体 ----
    if 'xl/metadata.xml' in zs.namelist() and 'xl/metadata.xml' not in parts:
        parts['xl/metadata.xml'] = zs.read('xl/metadata.xml')
        # 关系
        rk = 'xl/_rels/workbook.xml.rels'
        r = parts[rk].decode('utf8')
        if 'sheetMetadata' not in r:
            used = set(re.findall(r'Id="rId(\d+)"', r))
            nid = max((int(i) for i in used), default=0) + 1
            r = r.replace('</Relationships>',
                f'<Relationship Id="rId{nid}" Type="http://schemas.openxmlformats.org/'
                f'officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>'
                '</Relationships>')
            parts[rk] = r.encode('utf8')
        # 内容类型
        ck = '[Content_Types].xml'
        c = parts[ck].decode('utf8')
        if '/xl/metadata.xml' not in c:
            c = c.replace('</Types>',
                '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.'
                'openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>')
            parts[ck] = c.encode('utf8')
        done.append('metadata.xml')

    # ---- ③ 把 cm="1" 贴回对应的格子（按表名找，不按文件名）----
    n_cm = 0
    for nm, cells in want.items():
        path = smap_d.get(nm)
        if not path or path not in parts: continue
        x = parts[path].decode('utf8')
        for cell in cells:
            m = re.search(r'<c r="%s"(?![0-9])([^>]*)>' % cell, x)
            if not m or 'cm="1"' in m.group(0): continue
            # 只给「现在仍然是数组公式」的格子贴标记。
            # 往来款跟进 A11 原来是 =ROW(2:500) 的数组公式，这一版已经换成普通公式了，
            # 再贴 cm="1" 等于骗 Excel 说它是动态数组，反而会出问题。
            tail = x[m.end():m.end() + 60]
            if '<f t="array"' not in tail: continue
            x = x[:m.start()] + f'<c r="{cell}"{m.group(1)} cm="1">' + x[m.end():]
            n_cm += 1
        parts[path] = x.encode('utf8')
    if n_cm: done.append(f'{n_cm} 个 cm="1"')

    # ---- ④ 批注（comments + vml）----
    cmt = [n for n in zs.namelist() if n.startswith('xl/comments') or 'vmlDrawing' in n]
    if cmt and not any(n.startswith('xl/comments') for n in parts):
        # 原件里批注挂在哪张表
        host = None
        for nm, path in smap_s.items():
            rp = path.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
            if rp in zs.namelist() and b'comments' in zs.read(rp):
                host = nm; break
        tgt = smap_d.get(host) if host else None
        if tgt and tgt in parts:
            for n in cmt: parts[n] = zs.read(n)
            rp_d = tgt.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
            rels = parts.get(rp_d, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   b'</Relationships>').decode('utf8')
            used = set(int(i) for i in re.findall(r'Id="rId(\d+)"', rels))
            a = max(used, default=0) + 1
            if 'comments' not in rels:
                rels = rels.replace('</Relationships>',
                    f'<Relationship Id="rId{a}" Type="http://schemas.openxmlformats.org/'
                    f'officeDocument/2006/relationships/comments" Target="../comments1.xml"/>'
                    f'<Relationship Id="rId{a+1}" Type="http://schemas.openxmlformats.org/'
                    f'officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
                    '</Relationships>')
            parts[rp_d] = rels.encode('utf8')
            x = parts[tgt].decode('utf8')
            if '<legacyDrawing' not in x:
                x = re.sub(r'(</worksheet>)', f'<legacyDrawing r:id="rId{a+1}"/>\\1', x)
                parts[tgt] = x.encode('utf8')
            ck = '[Content_Types].xml'; c = parts[ck].decode('utf8')
            if '/xl/comments1.xml' not in c:
                c = c.replace('</Types>',
                    '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>')
            if 'vmlDrawing' not in c:
                c = c.replace('</Types>',
                    '<Default Extension="vml" ContentType="application/vnd.openxmlformats-'
                    'officedocument.vmlDrawing"/></Types>')
            parts[ck] = c.encode('utf8')
            done.append(f'【{host}】的批注')

    tmp = dst + '.fix'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, b in parts.items():
            if n.endswith('/'): continue
            zo.writestr(n, b)
    shutil.move(tmp, dst)
    zs.close()
    print('  补回：' + ('；'.join(done) if done else '（没有要补的）'))
    return done

if __name__ == '__main__':
    repair(sys.argv[1], sys.argv[2])
