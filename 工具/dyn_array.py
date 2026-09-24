# -*- coding: utf-8 -*-
"""给 openpyxl 存出来的工作簿补上「动态数组」标记。

openpyxl 会保留 <f t="array" ref="..."> 和 _xlfn._xlws.FILTER 的原文，
却把 xl/metadata.xml 和格子上的 cm="1" 一起扔掉。这两样是 Excel 365 / WPS
用来认「这是动态数组」的；标记一掉，公式退化成尺寸写死的老式 CSE 数组 ——
FILTER 只能铺满 ref 写死的那块区域，往下加数据不会自动长。

本脚本：
  · 没有 metadata.xml 就装一个（可以从 donor 文件原样拷，最保险）
  · 给每个「现在确实还是数组公式」的格子贴回 cm="1"

跑法：
  python3 dyn_array.py <xlsx> [donor.xlsx]
"""
import os, re, sys, shutil, zipfile

# 没有 donor 时用这份标准 XLDAPR 元数据
DEFAULT_META = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"'
    ' xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000"'
    ' copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1"'
    ' clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext>'
    '</extLst></bk></futureMetadata>'
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    '</metadata>'
).encode('utf8')

REL_META = ('http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/sheetMetadata')
CT_META = ('application/vnd.openxmlformats-officedocument.'
           'spreadsheetml.sheetMetadata+xml')


def _parts(path):
    z = zipfile.ZipFile(path)
    d = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    z.close()
    return d


def _write(path, parts):
    tmp = path + '.dyn'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, b in parts.items():
            zo.writestr(n, b)
    shutil.move(tmp, path)


def install(path, donor=None, verbose=True):
    parts = _parts(path)

    # ① metadata.xml 本体
    added_meta = False
    if 'xl/metadata.xml' not in parts:
        blob = DEFAULT_META
        if donor and os.path.exists(donor):
            dz = zipfile.ZipFile(donor)
            if 'xl/metadata.xml' in dz.namelist():
                blob = dz.read('xl/metadata.xml')
            dz.close()
        parts['xl/metadata.xml'] = blob
        added_meta = True

    # ② workbook 关系
    rk = 'xl/_rels/workbook.xml.rels'
    r = parts[rk].decode('utf8')
    if 'sheetMetadata' not in r:
        nid = max((int(i) for i in re.findall(r'Id="rId(\d+)"', r)), default=0) + 1
        r = r.replace('</Relationships>',
                      f'<Relationship Id="rId{nid}" Type="{REL_META}" '
                      f'Target="metadata.xml"/></Relationships>')
        parts[rk] = r.encode('utf8')

    # ③ 内容类型
    ck = '[Content_Types].xml'
    c = parts[ck].decode('utf8')
    if '/xl/metadata.xml' not in c:
        c = c.replace('</Types>',
                      f'<Override PartName="/xl/metadata.xml" '
                      f'ContentType="{CT_META}"/></Types>')
        parts[ck] = c.encode('utf8')

    # ④ 给还是数组公式的格子贴 cm="1"
    n_cm = 0
    hit = []
    for n in list(parts):
        if '/worksheets/' not in n or not n.endswith('.xml'):
            continue
        x = parts[n].decode('utf8')
        out, pos, changed = [], 0, False
        # 逐个 <c ...> 看**它自己**里面是不是 <f t="array"。
        # 只能看到下一个 <c 为止 —— 固定往后取 80 个字符会越界看到下一格的公式，
        # 于是把表头那种普通格子也当成数组公式贴上 cm="1"。
        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)>', x):
            nxt = x.find('<c ', m.end())
            tail = x[m.end():nxt if nxt != -1 else len(x)]
            if '<f t="array"' not in tail:
                continue
            if 'cm="' in m.group(2):
                continue
            out.append((m.start(), m.end(), m.group(1), m.group(2)))
            changed = True
        if not changed:
            continue
        buf, last = [], 0
        for s, e, cell, attrs in out:
            buf.append(x[last:s])
            buf.append(f'<c r="{cell}"{attrs} cm="1">')
            last = e
            n_cm += 1
            hit.append((n, cell))
        buf.append(x[last:])
        parts[n] = ''.join(buf).encode('utf8')

    _write(path, parts)
    if verbose:
        print('  %-34s %s，贴回 cm="1" %d 个 %s'
              % (os.path.basename(path)[:32],
                 '装上 metadata.xml' if added_meta else 'metadata.xml 已有',
                 n_cm, [c for _, c in hit[:8]]))
    return n_cm


if __name__ == '__main__':
    install(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
