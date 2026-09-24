# -*- coding: utf-8 -*-
"""扫工作簿里所有公式，挑出括号不配对 / 引号不配对的。

为什么要有这个：openpyxl 不校验公式，括号少写一个照样存得进去，
Excel / WPS 打开时会**静默把这格清空**（不报错、不提示），
于是整列「核对」公式在用户手里悄悄消失。LibreOffice 重算也可能直接丢掉，
所以只靠重算查不出来。

跑法：python3 check_formula.py <xlsx> [更多...]
"""
import re, sys, zipfile


def balance(f):
    """返回 (括号净差, 引号是否成对)。字符串里的括号不算。"""
    depth, i, n, inq = 0, 0, len(f), False
    while i < n:
        ch = f[i]
        if inq:
            if ch == '"':
                if i + 1 < n and f[i + 1] == '"':
                    i += 1
                else:
                    inq = False
        elif ch == '"':
            inq = True
        elif ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth < 0:
                return depth, True
        i += 1
    return depth, not inq


# 自闭合的 <c r="A1" s="3"/> 没有 </c>，所以不能写成 <c ...>(.*?)</c> ——
# 非贪婪会一路咬到后面某个格子的 </c>，报出来的坐标全是错的。
# 改成：切到下一个 <c 为止，再在这一段里找 <f>。
CELL = re.compile(r'<c r="([A-Z]+\d+)"')
FML = re.compile(r'<f[^>]*>(.*?)</f>', re.S)
UNESC = [('&lt;', '<'), ('&gt;', '>'), ('&amp;', '&'), ('&quot;', '"'), ('&apos;', "'")]


def scan(path, verbose=True):
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode('utf8')
    names = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="rId(\d+)"', wbx)
    rels = {}
    for tag in re.findall(r'<Relationship [^>]*/>',
                          z.read('xl/_rels/workbook.xml.rels').decode('utf8')):
        a = dict(re.findall(r'([A-Za-z]+)="([^"]*)"', tag))
        if a.get('Id', '').startswith('rId'):
            rels[a['Id'][3:]] = a.get('Target', '')
    sheet_of = {}
    for nm, rid in names:
        t = rels.get(rid, '')
        sheet_of['xl/' + t.lstrip('/')] = nm
    bad, total = [], 0
    for n in z.namelist():
        if '/worksheets/' not in n or not n.endswith('.xml'):
            continue
        sn = sheet_of.get(n, n)
        x = z.read(n).decode('utf8')
        for m in CELL.finditer(x):
            nxt = x.find('<c ', m.end())
            fm = FML.search(x[m.end():nxt if nxt != -1 else len(x)])
            if not fm:
                continue
            f = fm.group(1)
            for a, b in UNESC:
                f = f.replace(a, b)
            total += 1
            d, q = balance(f)
            if d != 0 or not q:
                bad.append((sn, m.group(1), d, q, f[:150]))
    z.close()
    if verbose:
        print('%-46s 公式 %6d 条，坏 %d 条' % (path.split('/')[-1][:44], total, len(bad)))
        seen = set()
        for sn, cell, d, q, f in bad:
            key = (sn, re.sub(r'\d+', '#', f[:80]))
            if key in seen:
                continue
            seen.add(key)
            print('   ✗ %s!%s  括号差 %+d  引号%s' % (sn, cell, d, '成对' if q else '不成对'))
            print('     %s' % f)
    return bad


if __name__ == '__main__':
    tot = 0
    for p in sys.argv[1:]:
        tot += len(scan(p))
    sys.exit(1 if tot else 0)
