# -*- coding: utf-8 -*-
"""稳妥地跑一次 LibreOffice 重算，保证「一定会返回」。

之前踩的三个坑，都在这里堵掉了：

1. **残留的 .~lock.xxx.xlsx#**
   LibreOffice 被强杀后会把锁文件留在原地。下一次重算它既不报错也不干活，
   recalc.py 就一直不输出 —— 外面等结果的循环于是永远转下去。
   开工前必须先删锁。

2. **`rm -f .~lock.xxx.xlsx#` 删不掉**
   这个壳里 `#` 的处理跟 bash 不一样，那行 rm 实际执行的是
   `rm -f .~lock.xxx.xlsx`（`#` 后面被当注释吃掉了），文件纹丝不动。
   所以只能用 Python 的 os.remove，别再用 shell。

3. **`pkill -f soffice` 会把自己打死**
   跑这条命令的 bash，它自己的命令行里就含 "soffice" 四个字母，
   pkill -f 一匹配就把当前这个 shell 也杀了 —— 表现为「命令毫无输出」，
   看着像卡住，其实是自杀了。
   所以清理旧进程只能按 PID 来，而且要跳过自己和自己的祖先。

跑法：python3 recalc_safe.py <xlsx> [超时秒数] [--baseline N]
      --baseline N ：这本册子原来就有 N 处报错（比如原作者用了 LibreOffice
                     不认的函数），不超过 N 就算通过。
返回码：0 = 通过，1 = 有新增报错，2 = 重算本身没跑成
"""
import os, sys, glob, json, signal, shutil, subprocess

RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'


def _ancestors(pid):
    """自己和自己的所有祖先，这些 PID 一个都不能杀"""
    out, seen = set(), pid
    while seen and seen not in out:
        out.add(seen)
        try:
            with open(f'/proc/{seen}/stat') as f:
                seen = int(f.read().split(') ', 1)[1].split()[1])
        except Exception:
            break
    return out


def kill_stale_soffice(verbose=True):
    """按 PID 杀掉遗留的 LibreOffice，绝不用 pkill -f（会误杀自己）"""
    safe = _ancestors(os.getpid())
    killed = []
    try:
        ps = subprocess.run(['ps', '-eo', 'pid,args'], capture_output=True, text=True).stdout
    except Exception:
        return killed
    for line in ps.splitlines()[1:]:
        line = line.strip()
        if not line:
            continue
        head, _, args = line.partition(' ')
        if not head.isdigit():
            continue
        pid = int(head)
        if pid in safe:
            continue
        # 只认真正的 LibreOffice 进程，不认那些「命令行里碰巧提到 soffice」的 shell
        if ('/program/soffice' in args or args.startswith('soffice')
                or 'oosplash' in args or 'RecalculateAndSave' in args):
            try:
                os.kill(pid, signal.SIGKILL)
                killed.append(pid)
            except OSError:
                pass
    if killed and verbose:
        print(f'  清掉遗留的 LibreOffice 进程 {len(killed)} 个：{killed}')
    return killed


def clear_locks(path, verbose=True):
    """删掉这个文件旁边的锁，以及所有临时配置目录"""
    d = os.path.dirname(os.path.abspath(path)) or '.'
    base = os.path.basename(path)
    gone = []
    for lk in [os.path.join(d, f'.~lock.{base}#')] + glob.glob(os.path.join(d, '.~lock.*#')):
        if os.path.exists(lk):
            try:
                os.remove(lk); gone.append(os.path.basename(lk))
            except OSError:
                pass
    for prof in glob.glob('/tmp/recalc-lo-profile-*'):
        shutil.rmtree(prof, ignore_errors=True)
    if gone and verbose:
        print(f'  删掉残留锁文件：{gone}')
    return gone


def recalc(path, timeout=900, baseline=0, verbose=True):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        print(f'✗ 文件不在：{path}'); return 2, None
    kill_stale_soffice(verbose)
    clear_locks(path, verbose)

    # 给外层留 60 秒余量，免得 recalc.py 自己还没写完结果就被砍
    hard = timeout + 60
    try:
        r = subprocess.run([sys.executable, RECALC, path, str(timeout), '--force'],
                           capture_output=True, text=True, timeout=hard)
        out = r.stdout
    except subprocess.TimeoutExpired:
        print(f'✗ 重算超过 {hard} 秒还没结束，已强制中断')
        kill_stale_soffice(verbose); clear_locks(path, verbose)
        return 2, None

    i, j = out.find('{'), out.rfind('}')
    if i < 0 or j < 0:
        print('✗ 重算没有返回 JSON，原始输出：')
        print((out or r.stderr or '（空）')[:600])
        kill_stale_soffice(verbose); clear_locks(path, verbose)
        return 2, None
    try:
        st = json.loads(out[i:j + 1])
    except json.JSONDecodeError as e:
        print('✗ 返回的 JSON 解析不了：', e); return 2, None

    clear_locks(path, verbose=False)          # 正常结束也可能留锁，顺手清掉
    errs, nf = st.get('total_errors', 0), st.get('total_formulas', 0)
    tag = '✓' if errs <= baseline else '✗'
    extra = f'（原表基线 {baseline}）' if baseline else ''
    print(f'{tag} 重算完成：公式 {nf} 个，报错 {errs} 处{extra}')
    for k, v in (st.get('error_summary') or {}).items():
        print(f'    {k} ×{v["count"]}  {v["locations"][:6]}')
    return (0 if errs <= baseline else 1), st


if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    base = 0
    if '--baseline' in sys.argv:
        base = int(sys.argv[sys.argv.index('--baseline') + 1])
    code, _ = recalc(a[0], int(a[1]) if len(a) > 1 else 900, base)
    sys.exit(code)
