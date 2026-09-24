# -*- coding: utf-8 -*-
"""给工作簿加一条「引用另一本工作簿」的跨文件链接。

openpyxl 建不出 externalLink 部件，所以这里直接按 OOXML 拼：
  xl/externalLinks/externalLinkN.xml            —— 记住对方有哪几张表
  xl/externalLinks/_rels/externalLinkN.xml.rels —— TargetMode="External" 指向对方文件名
  xl/workbook.xml 里加 <externalReference r:id="..."/>（必须排在 <sheets> 之后、
                                                        <definedNames> 之前，顺序错了 Excel 会报修复）
  [Content_Types].xml 加 Override

加完之后，公式里就可以写  [1]数据录入!$A$19:$K$5078  —— [1] 是本函数返回的序号。

注意：没有缓存值。对方文件**关着**的时候这些公式取不到数（FILTER/SUMIFS 这类
本来就不支持读关闭的工作簿）。所以两本必须同时打开 —— 这正是用户要的联动方式。

跑法（当模块用）：
  from ext_link import add
  idx = add('本册.xlsx', '对方.xlsx', ['数据录入', '基础资料'])
"""
import re, shutil, zipfile

REL_EXT = ('http://schemas.openxmlformats.org/officeDocument/2006/'
           'relationships/externalLink')
REL_PATH = ('http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/externalLinkPath')
CT_EXT = ('application/vnd.openxmlformats-officedocument.'
          'spreadsheetml.externalLink+xml')
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def add(path, target_file, sheet_names, verbose=True):
    z = zipfile.ZipFile(path)
    parts = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    z.close()

    exist = [n for n in parts if re.match(r'xl/externalLinks/externalLink\d+\.xml$', n)]
    idx = len(exist) + 1

    names = ''.join(f'<sheetName val="{s}"/>' for s in sheet_names)
    data = ''.join(f'<sheetData sheetId="{i}"/>' for i in range(len(sheet_names)))
    parts[f'xl/externalLinks/externalLink{idx}.xml'] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<externalLink xmlns="{NS}"><externalBook xmlns:r="{NSR}" r:id="rId1">'
        f'<sheetNames>{names}</sheetNames>'   # 不能加 count：ECMA-376 里 sheetNames 没这个属性
        f'<sheetDataSet>{data}</sheetDataSet>'
        f'</externalBook></externalLink>').encode('utf8')

    parts[f'xl/externalLinks/_rels/externalLink{idx}.xml.rels'] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL_PATH}" Target="{target_file}" '
        f'TargetMode="External"/></Relationships>').encode('utf8')

    rk = 'xl/_rels/workbook.xml.rels'
    r = parts[rk].decode('utf8')
    nid = max((int(i) for i in re.findall(r'Id="rId(\d+)"', r)), default=0) + 1
    rid = f'rId{nid}'
    r = r.replace('</Relationships>',
                  f'<Relationship Id="{rid}" Type="{REL_EXT}" '
                  f'Target="externalLinks/externalLink{idx}.xml"/></Relationships>')
    parts[rk] = r.encode('utf8')

    wb = parts['xl/workbook.xml'].decode('utf8')
    if '<externalReferences>' in wb:
        wb = wb.replace('</externalReferences>',
                        f'<externalReference r:id="{rid}"/></externalReferences>')
    else:
        block = f'<externalReferences><externalReference r:id="{rid}"/></externalReferences>'
        if '<definedNames>' in wb:          # 必须插在 definedNames 前面
            wb = wb.replace('<definedNames>', block + '<definedNames>', 1)
        elif '</sheets>' in wb:
            wb = wb.replace('</sheets>', '</sheets>' + block, 1)
        else:
            raise RuntimeError('workbook.xml 里既没有 sheets 也没有 definedNames，插不进去')
    parts['xl/workbook.xml'] = wb.encode('utf8')

    ck = '[Content_Types].xml'
    c = parts[ck].decode('utf8')
    tag = f'/xl/externalLinks/externalLink{idx}.xml'
    if tag not in c:
        c = c.replace('</Types>',
                      f'<Override PartName="{tag}" ContentType="{CT_EXT}"/></Types>')
        parts[ck] = c.encode('utf8')

    tmp = path + '.ext'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, b in parts.items():
            zo.writestr(n, b)
    shutil.move(tmp, path)
    if verbose:
        print(f'  加跨文件链接 [{idx}] → {target_file}  （表：{"/".join(sheet_names)}）')
    return idx
