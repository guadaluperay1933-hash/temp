# -*- coding: utf-8 -*-
"""一键把四本台账改造好：原件在 ../参考/，成品写到 ../ 下（文件名必须保持不变，
   跨文件链接是按文件名找的）。

跑法：python3 build_all.py [输出目录]
"""
import os, sys, shutil, subprocess, time

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, '..', '参考')
OUTDIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..')
TMP = os.path.join(OUTDIR, '_tmp')
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
NAMES = {'01': '01_水果进销存台账模板.xlsx', '02': '02_物料与周转物台账模板.xlsx',
         '03': '03_财务账套与报表模板.xlsx', '04': '04_综合查询对账单模板.xlsx'}

def run(*a):
    print('  $', ' '.join(os.path.basename(x) for x in a))
    r = subprocess.run([sys.executable] + list(a), capture_output=True, text=True)
    if r.returncode: print(r.stdout, r.stderr); raise SystemExit('失败: ' + a[0])
    for line in r.stdout.splitlines(): print('   ', line)

def recalc(path, timeout=1500, force=True):
    args = [sys.executable, RECALC, path, str(timeout)] + (['--force'] if force else [])
    r = subprocess.run(args, capture_output=True, text=True)
    print('   重算', os.path.basename(path), r.stdout.strip().replace('\n', ' ')[:200])

os.makedirs(TMP, exist_ok=True)
S = lambda n: os.path.join(TMP, n)

print('① 改《01 水果进销存台账》')
run(os.path.join(HERE, 'fix01_购买方.py'), os.path.join(REF, '原01_水果进销存台账模板.xlsx'), S('a01.xlsx'))
run(os.path.join(HERE, 'fix02_接口.py'), S('a01.xlsx'), S('b01.xlsx'))
run(os.path.join(HERE, 'fix03_新表.py'), S('b01.xlsx'), os.path.join(OUTDIR, NAMES['01']))

print('② 《02 物料与周转物台账》这一轮没有要改的，原样带过去')
shutil.copy(os.path.join(REF, '原02_物料与周转物台账模板.xlsx'), os.path.join(OUTDIR, NAMES['02']))

print('③ 改《03 财务账套与报表》')
run(os.path.join(HERE, 'fix04_03表.py'), os.path.join(REF, '原03_财务账套与报表模板.xlsx'), S('a03.xlsx'))
run(os.path.join(HERE, 'repair_ext_links.py'), S('a03.xlsx'), os.path.join(REF, '原03_财务账套与报表模板.xlsx'))
run(os.path.join(HERE, 'fix06_03新表.py'), S('a03.xlsx'), os.path.join(OUTDIR, NAMES['03']))
run(os.path.join(HERE, 'repair_ext_links.py'), os.path.join(OUTDIR, NAMES['03']), S('a03.xlsx'))

print('④ 改《04 综合查询对账单》')
run(os.path.join(HERE, 'fix05_04表.py'), os.path.join(REF, '原04_综合查询对账单模板.xlsx'), os.path.join(OUTDIR, NAMES['04']))
run(os.path.join(HERE, 'repair_ext_links.py'), os.path.join(OUTDIR, NAMES['04']), os.path.join(REF, '原04_综合查询对账单模板.xlsx'))

print('⑤ 三本册子的主页补上新表入口')
run(os.path.join(HERE, 'fix07_主页.py'), OUTDIR)
for k in ('03', '04'):
    run(os.path.join(HERE, 'repair_ext_links.py'), os.path.join(OUTDIR, NAMES[k]),
        os.path.join(REF, '原' + NAMES[k]))

print('⑥ 按依赖顺序重算 + 刷新跨文件缓存')
p = lambda k: os.path.join(OUTDIR, NAMES[k])
recalc(p('01')); recalc(p('02'))
for k in ('03', '04'):
    run(os.path.join(HERE, 'refresh_link_cache.py'), p(k), p('01'), p('02'))
    recalc(p(k))
    run(os.path.join(HERE, 'refresh_link_cache.py'), p(k), p('01'), p('02'))
    recalc(p(k))
shutil.rmtree(TMP, ignore_errors=True)
print('\n全部完成，成品在', os.path.abspath(OUTDIR))
