# -*- coding: utf-8 -*-
"""给《资金台账》的「数据录入」加"录入即出边框"，并改名成 A061_资金台账.xlsx。

为什么不用 openpyxl 存一遍：这本是买来的模板，里头有 14 条 FILTER 动态数组、
两组批注、richData、一个 Excel 表格对象。openpyxl 存盘会把 metadata.xml 和
cm="1" 一起扔掉（动态数组退化），richData 也没了。所以直接改 XML —— 只往
sheet3 里插一段条件格式、往 styles.xml 里加一个 dxf，别的一个字节不动。

边框是**条件格式**做的，没有宏：B 列（日期）一填，整行 A:K 的框线自己出来，
删掉又自己消失。WPS 一样认。
"""
import os, re, shutil, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, '参考', '原_资金台账.xlsx')
OUT = os.path.join(ROOT, 'A061_资金台账.xlsx')

SHEET = '数据录入'
SQREF = 'A5:K5078'          # 跟这张表公式的设计范围一致
COND = '$B5&lt;&gt;""'      # B 列有日期就算这一行有数据

DXF = ('<dxf><border>'
       '<left style="thin"><color rgb="FF808080"/></left>'
       '<right style="thin"><color rgb="FF808080"/></right>'
       '<top style="thin"><color rgb="FF808080"/></top>'
       '<bottom style="thin"><color rgb="FF808080"/></bottom>'
       '</border></dxf>')


def _attrs(tag):
    return dict(re.findall(r'([A-Za-z:]+)="([^"]*)"', tag))


def sheet_path(parts, name):
    wb = parts['xl/workbook.xml'].decode('utf8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf8')
    rid2t = {}
    for m in re.finditer(r'<Relationship[^>]*/?>', rels):
        a = _attrs(m.group(0))
        if a.get('Id') and a.get('Target'):
            t = a['Target'].lstrip('/')
            rid2t[a['Id']] = t if t.startswith('xl/') else 'xl/' + t
    for m in re.finditer(r'<sheet[^>]*/?>', wb):
        a = _attrs(m.group(0))
        if a.get('name') == name:
            return rid2t.get(a.get('r:id', ''))
    return None


def main():
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    z.close()

    # ① styles.xml 里加一个只有边框的 dxf，记下它的序号
    st = parts['xl/styles.xml'].decode('utf8')
    m = re.search(r'<dxfs count="(\d+)">', st)
    if m:
        n = int(m.group(1))
        st = st[:m.start()] + f'<dxfs count="{n + 1}">' + st[m.end():]
        st = st.replace('</dxfs>', DXF + '</dxfs>', 1)
    elif '<dxfs count="0"/>' in st:
        n = 0
        st = st.replace('<dxfs count="0"/>', f'<dxfs count="1">{DXF}</dxfs>', 1)
    else:
        n = 0
        st = st.replace('</styleSheet>', f'<dxfs count="1">{DXF}</dxfs></styleSheet>', 1)
    parts['xl/styles.xml'] = st.encode('utf8')

    # ② sheet 里插条件格式。位置有讲究：必须排在 dataValidations 之前，
    #    不然 Excel 会判成结构错误弹"需要修复"。
    p = sheet_path(parts, SHEET)
    if not p or p not in parts:
        sys.exit('找不到工作表 ' + SHEET)
    x = parts[p].decode('utf8')
    if f'sqref="{SQREF}"' in x and 'cfRule' in x:
        print('  已经加过了，跳过')
    else:
        cf = (f'<conditionalFormatting sqref="{SQREF}">'
              f'<cfRule type="expression" dxfId="{n}" priority="1">'
              f'<formula>{COND}</formula></cfRule></conditionalFormatting>')
        for anchor in ('<dataValidations', '<pageMargins', '<headerFooter', '</worksheet>'):
            i = x.find(anchor)
            if i != -1:
                x = x[:i] + cf + x[i:]
                break
        else:
            sys.exit('sheet3 里找不到可以插入的位置')
        parts[p] = x.encode('utf8')

    tmp = OUT + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in parts.items():
            zo.writestr(k, b)
    shutil.move(tmp, OUT)
    print(f'  已生成 {os.path.basename(OUT)}：'
          f'【{SHEET}】{SQREF} 加了"有数据就出边框"的条件格式（dxfId={n}）')


if __name__ == '__main__':
    main()
