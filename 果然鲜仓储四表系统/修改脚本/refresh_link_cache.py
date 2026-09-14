# -*- coding: utf-8 -*-
"""把《03》《04》里「跨文件链接的缓存值」按最新的《01》《02》重写一遍。

Excel 打开时会问你要不要更新链接，更新了就是最新的；但没更新之前显示的是文件里存着的
那份缓存。改完 01 之后如果不把缓存一起刷新，打开 03/04 第一眼看到的还是旧数。
本脚本就是把这份缓存重新算好写进去。

跑法：python3 refresh_link_cache.py <目标xlsx> <源1.xlsx> <源2.xlsx> ...
      源文件按 externalLink1、2… 的顺序给，或者用「文件名=路径」显式指定。
"""
import sys, os, re, zipfile, shutil, datetime as dt
import openpyxl
from openpyxl.utils import column_index_from_string as CI

REF = re.compile(r"\[(\d+)\](?:'([^']+)'|([^!'\[\]\s,()+\-*/&=<>]+))!(\$?[A-Z]{1,3}\$?\d+)")

def norm(c):
    return c.replace('$', '')

def xml_escape(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))

def collect_refs(path):
    """扫目标工作簿的全部公式，收集 [n]表名!单元格"""
    wb = openpyxl.load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    for m in REF.finditer(c.value):
                        idx = int(m.group(1))
                        sheet = m.group(2) or m.group(3)
                        out.setdefault(idx, {}).setdefault(sheet, set()).add(norm(m.group(4)))
    wb.close()
    return out

def link_targets(path):
    """externalLink 序号 → 被链接的文件名"""
    z = zipfile.ZipFile(path)
    wbxml = z.read('xl/workbook.xml').decode('utf8')
    m = re.search(r'<externalReferences>(.*?)</externalReferences>', wbxml, re.S)
    ids = re.findall(r'r:id="(rId\d+)"', m.group(1)) if m else []
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rel = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for k, rid in enumerate(ids, 1):
        part = rel[rid].replace('../', '')
        nm = part.split('/')[-1]
        tgt = re.search(rb'Target="([^"]+)"',
                        z.read('xl/externalLinks/_rels/' + nm + '.rels')).group(1).decode()
        out[k] = (nm, tgt)
    z.close()
    return out

def build_sheetdata(src_path, sheets_cells):
    """从源文件读值，拼出 <sheetNames> 与 <sheetDataSet>"""
    wb = openpyxl.load_workbook(src_path, data_only=True)
    names = wb.sheetnames
    parts = []
    for sid, nm in enumerate(names):
        cells = sheets_cells.get(nm)
        if not cells:
            parts.append(f'<sheetData sheetId="{sid}" refreshError="1"/>')
            continue
        ws = wb[nm]
        byrow = {}
        for a in cells:
            m = re.match(r'([A-Z]+)(\d+)$', a)
            byrow.setdefault(int(m.group(2)), []).append((CI(m.group(1)), a))
        rows = []
        for r in sorted(byrow):
            cs = []
            for _, a in sorted(byrow[r]):
                v = ws[a].value
                if v is None or v == '':
                    cs.append(f'<cell r="{a}" t="str"><v></v></cell>')
                elif isinstance(v, bool):
                    cs.append(f'<cell r="{a}" t="b"><v>{1 if v else 0}</v></cell>')
                elif isinstance(v, (int, float)):
                    cs.append(f'<cell r="{a}"><v>{v!r}</v></cell>')
                elif isinstance(v, (dt.datetime, dt.date)):
                    d0 = dt.datetime(1899, 12, 30)
                    d = v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
                    cs.append(f'<cell r="{a}"><v>{(d - d0).days + (d - d0).seconds / 86400.0!r}</v></cell>')
                else:
                    cs.append(f'<cell r="{a}" t="str"><v>{xml_escape(v)}</v></cell>')
            rows.append(f'<row r="{r}">{"".join(cs)}</row>')
        parts.append(f'<sheetData sheetId="{sid}">{"".join(rows)}</sheetData>')
    wb.close()
    sn = ''.join(f'<sheetName val="{xml_escape(n)}"/>' for n in names)
    return f'<sheetNames>{sn}</sheetNames>', f'<sheetDataSet>{"".join(parts)}</sheetDataSet>'

def refresh(target, srcmap):
    refs = collect_refs(target)
    links = link_targets(target)
    newparts = {}
    for idx, (partname, tgtfile) in links.items():
        src = srcmap.get(tgtfile)
        if not src:
            print(f'   [{idx}] {tgtfile} 没给源文件，跳过'); continue
        sn, sd = build_sheetdata(src, refs.get(idx, {}))
        newparts['xl/externalLinks/' + partname] = (sn, sd)
        ncell = sum(len(v) for v in refs.get(idx, {}).values())
        print(f'   [{idx}] {tgtfile} ← {os.path.basename(src)}：刷新 {ncell} 个格子的缓存')
    tmp = target + '.tmp'
    zin = zipfile.ZipFile(target); zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename in newparts:
            t = data.decode('utf8')
            sn, sd = newparts[it.filename]
            t = re.sub(r'<sheetNames>.*?</sheetNames>', sn, t, flags=re.S)
            if '<sheetDataSet>' in t or '<sheetDataSet/>' in t:
                t = re.sub(r'<sheetDataSet>.*?</sheetDataSet>|<sheetDataSet/>', sd, t, flags=re.S)
            else:
                t = t.replace('</externalBook>', sd + '</externalBook>')
            data = t.encode('utf8')
        zout.writestr(it, data)
    zin.close(); zout.close()
    shutil.move(tmp, target)

if __name__ == '__main__':
    target = sys.argv[1]
    srcmap = {}
    for a in sys.argv[2:]:
        if '=' in a:
            k, v = a.split('=', 1); srcmap[k] = v
        else:
            srcmap[os.path.basename(a)] = a
    print('刷新', os.path.basename(target))
    refresh(target, srcmap)
    print('   完成')
