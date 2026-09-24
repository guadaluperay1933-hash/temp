# -*- coding: utf-8 -*-
"""把工作簿的「当前选中哪张表」修正成正常状态。

为什么要修：
Excel / WPS 认为一本工作簿里应该**正好有一张**表是选中的（sheetView 上的
tabSelected="1"）。openpyxl 存盘出来经常是**一张都没有**，只在 workbook.xml 里留个
activeTab。WPS 碰到这种状态会判定成「选择了多个工作表」（工作组），于是：
  · 弹「该工作簿中包含一个或多个无法更新的链接。并且在上次保存该工作簿时，
    选择了多个工作表。在这种状态下无法修复链接。」
  · 【数据 → 编辑链接】整个变灰，点不动，连手工重新指定源文件都做不了。

修法：让 activeTab 指的那张表（越界就取第一张可见表）带上 tabSelected="1"，
其余所有表清掉。只改 sheetView 这一个属性，别的部件一个字节都不动 ——
跨文件链接、动态数组标记、缓存值全部原样保留。

跑法：python3 fix_sheet_selection.py <xlsx> [更多 xlsx...]
"""
import os, re, sys, shutil, zipfile


def _attrs(tag):
    return dict(re.findall(r'([A-Za-z:]+)="([^"]*)"', tag))


def _norm(t):
    t = t.lstrip('/')
    return t if t.startswith('xl/') else 'xl/' + t


SV = r'<sheetView(?![A-Za-z])[^>]*?/?>'      # (?![A-Za-z]) 防止咬到容器标签 <sheetViews>
BROKEN = re.compile(r'<sheetView((?:\s+[A-Za-z:]+="[^"]*")*)s>')


def unbreak(x):
    """把本脚本早期版本写坏的 <sheetView tabSelected="1"s> 还原成 <sheetViews>。"""
    return BROKEN.sub('<sheetViews>', x)


def fix(path, verbose=True):
    z = zipfile.ZipFile(path)
    parts = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    z.close()
    for n in list(parts):
        if '/worksheets/' in n and n.endswith('.xml'):
            t = parts[n].decode('utf8')
            t2 = unbreak(t)
            if t2 != t:
                parts[n] = t2.encode('utf8')

    wb = parts['xl/workbook.xml'].decode('utf8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf8')
    rid2t = {}
    for m in re.finditer(r'<Relationship[^>]*/?>', rels):
        a = _attrs(m.group(0))
        if a.get('Id') and a.get('Target'):
            rid2t[a['Id']] = _norm(a['Target'])

    sheets = []                       # [(表名, 部件路径, 是否隐藏)]
    for m in re.finditer(r'<sheet[^>]*/?>', wb):
        a = _attrs(m.group(0))
        p = rid2t.get(a.get('r:id', ''))
        if p and '/worksheets/' in p:
            sheets.append((a.get('name', ''), p, a.get('state', 'visible') != 'visible'))
    if not sheets:
        print(f'  {os.path.basename(path)}：没找到工作表，跳过'); return False

    m = re.search(r'activeTab="(\d+)"', wb)
    # 模板打开时该停在【主页】，而不是原作者最后停留的那张表（03 原来停在「科目余额表」）
    idx = next((i for i, s in enumerate(sheets)
                if s[0] in ('主页', '首页') and not s[2]), None)
    if idx is None:
        idx = int(m.group(1)) if m else 0
        if idx >= len(sheets) or sheets[idx][2]:      # 越界或指到隐藏表
            idx = next((i for i, s in enumerate(sheets) if not s[2]), 0)

    before = [nm for nm, p, _ in sheets
              if 'tabSelected="1"' in (re.search(r'<sheetView[^>]*>', parts[p].decode('utf8', 'ignore'))
                                       or re.match('', '')).group(0)] if False else []
    before = []
    for nm, p, _ in sheets:
        sv = re.search(SV, parts[p].decode('utf8', 'ignore'))
        if sv and 'tabSelected="1"' in sv.group(0):
            before.append(nm)

    changed = 0
    for i, (nm, p, _) in enumerate(sheets):
        x = parts[p].decode('utf8')
        sv = re.search(SV, x)
        if not sv:
            continue
        tag = sv.group(0)
        new = re.sub(r'\s*tabSelected="[^"]*"', '', tag)      # 先一律清掉
        if i == idx:                                          # 该选中的那张补上
            new = re.sub(r'(<sheetView)', r'\1 tabSelected="1"', new, count=1)
        if new != tag:
            parts[p] = (x[:sv.start()] + new + x[sv.end():]).encode('utf8')
            changed += 1

    # activeTab 也校正一下，指到我们选中的那张
    if m:
        wb2 = wb[:m.start()] + f'activeTab="{idx}"' + wb[m.end():]
    else:
        wb2 = re.sub(r'(<workbookView\b[^>]*?)(/?>)', rf'\1 activeTab="{idx}"\2', wb, count=1)
    if wb2 != wb:
        parts['xl/workbook.xml'] = wb2.encode('utf8')

    tmp = path + '.sel'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, b in parts.items():
            zo.writestr(n, b)
    shutil.move(tmp, path)
    if verbose:
        print('  %-38s 原来选中 %d 张%s → 现在只选中【%s】，改了 %d 张表'
              % (os.path.basename(path)[:36], len(before),
                 ('（' + '/'.join(before[:4]) + '）') if before else '（一张都没选，正是这个毛病）',
                 sheets[idx][0], changed))
    return True


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fix(p)
