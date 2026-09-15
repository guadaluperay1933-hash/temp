# -*- coding: utf-8 -*-
"""openpyxl 存盘时会把 workbook.xml.rels 里的 externalLink 关系丢掉，
   只剩 workbook.xml 里那几个 <externalReference r:id="..."/>，指向的关系已经不存在了 ——
   这样的文件 Excel 打开会把跨文件链接当坏链处理。本脚本把这几条关系按原文件补回去。

跑法：python3 repair_ext_links.py <被 openpyxl 存过的 xlsx> <原始 xlsx>
"""
import sys, re, zipfile, shutil, os

def parts(path):
    z = zipfile.ZipFile(path)
    d = {n: z.read(n) for n in z.namelist()}
    infos = {i.filename: i for i in z.infolist()}
    z.close()
    return d, infos

def ext_map(xml_wb, xml_rels):
    m = re.search(r'<externalReferences>(.*?)</externalReferences>', xml_wb, re.S)
    if not m: return [], {}
    ids = re.findall(r'r:id="(rId\d+)"', m.group(1))
    rel = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', xml_rels))
    return ids, {i: rel[i] for i in ids if i in rel}

def repair(fixed, orig):
    od, _ = parts(orig)
    fd, finfos = parts(fixed)
    owb = od['xl/workbook.xml'].decode('utf8')
    orels = od['xl/_rels/workbook.xml.rels'].decode('utf8')
    ids, mp = ext_map(owb, orels)
    if not ids:
        print('   原文件没有跨文件链接，不用修'); return False
    targets = [mp[i] for i in ids if i in mp]
    fwb = fd['xl/workbook.xml'].decode('utf8')
    frels = fd['xl/_rels/workbook.xml.rels'].decode('utf8')
    used = set(re.findall(r'Id="(rId\d+)"', frels))
    n = 1
    newids = []
    for t in targets:
        while f'rId{n}' in used: n += 1
        newids.append(f'rId{n}'); used.add(f'rId{n}')
    # ① 补关系
    add = ''.join(
        f'<Relationship Id="{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/externalLink" Target="{t}"/>' for i, t in zip(newids, targets))
    if 'externalLink' in frels:
        print('   关系还在，只对齐编号')
        frels = re.sub(r'<Relationship[^>]*externalLink[^>]*/>', '', frels)
    frels = frels.replace('</Relationships>', add + '</Relationships>')
    # ② workbook.xml 里的 externalReferences 用新编号
    newrefs = ''.join(f'<externalReference r:id="{i}"/>' for i in newids)
    fwb = re.sub(r'<externalReferences>.*?</externalReferences>',
                 f'<externalReferences>{newrefs}</externalReferences>', fwb, flags=re.S)
    fd['xl/workbook.xml'] = fwb.encode('utf8')
    fd['xl/_rels/workbook.xml.rels'] = frels.encode('utf8')
    # ③ externalLink 的 xml 与它自己的 rels：只补「缺了的」，绝对不能拿原文件的覆盖。
    #    LibreOffice 存盘时会调整 externalReference 的先后（公式里的 [1]/[2] 跟着变），
    #    这时每个 externalLinkN.xml 指向哪个文件是由它自己的 .rels 说了算的；
    #    一旦拿原文件的 .rels 盖上去，[1]/[2] 就跟实际文件对调了 —— 对接源会读到另一本册子。
    for n2 in od:
        if n2.startswith('xl/externalLinks/') and n2 not in fd:
            fd[n2] = od[n2]
    # ④ Content_Types 里补声明
    ct = fd['[Content_Types].xml'].decode('utf8')
    for n2 in sorted(x for x in fd if re.match(r'xl/externalLinks/externalLink\d+\.xml$', x)):
        if f'/{n2}"' not in ct:
            ct = ct.replace('</Types>',
                            f'<Override PartName="/{n2}" ContentType="application/vnd.openxmlformats-'
                            f'officedocument.spreadsheetml.externalLink+xml"/></Types>')
    fd['[Content_Types].xml'] = ct.encode('utf8')
    tmp = fixed + '.tmp'
    zo = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for k, v in fd.items(): zo.writestr(k, v)
    zo.close(); shutil.move(tmp, fixed)
    print(f'   补回 {len(newids)} 条 externalLink 关系：{list(zip(newids, targets))}')
    return True

if __name__ == '__main__':
    print('修', os.path.basename(sys.argv[1]))
    repair(sys.argv[1], sys.argv[2])
