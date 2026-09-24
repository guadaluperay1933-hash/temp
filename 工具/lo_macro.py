# -*- coding: utf-8 -*-
"""在 LibreOffice 里对一本 xlsx 跑一段 Basic 宏（改完、重算、另存），用来模拟用户的手工操作。

为什么要有这个：公式对不对，光看公式文本看不出「用户删一行 / 插一行 / 排个序以后还对不对」。
Excel / WPS 删行时会把指向被删格子的引用改成 #REF!，插行时会把引用往下挪 ——
这些只有真在表格软件里做一遍才看得到。LibreOffice 处理引用的规则跟 Excel 一致，拿它来模拟。

用法（命令行）：
  python3 lo_macro.py <输入.xlsx> <输出.xlsx> <宏文件.bas> [超时秒数]

宏文件里写 Basic 语句，可以直接用 Doc（= 打开的工作簿）和 Sh(名字)（= 取工作表），例如：
  Sh("成品出库明细").Rows.removeByIndex(50, 4)     ' 删掉第 51~54 行（0 起算）
  Sh("成品出库明细").Rows.insertByIndex(10, 2)     ' 在第 11 行前插 2 行
  Sh("成品出库明细").getCellRangeByName("B80").setValue(46300)
  Sh("成品出库明细").getCellRangeByName("D80").setString("苹果")
跑完会自动 calculateAll 并另存成 xlsx（原文件不动）。

也可以在 Python 里 import：run_macro(src, dst, body, timeout)

注意：
- 输出是 LibreOffice 存的 xlsx，只拿来读数验证（openpyxl data_only=True），**不要当交付件**
  （LibreOffice 存盘会改写一些东西：动态数组标记、函数大小写、跨文件链接关系等）。
- 同一时间只跑一个 soffice 最稳；每次用独立的临时 profile，互不干扰。
"""
import os, sys, shutil, subprocess, tempfile, time
from pathlib import Path

sys.path.insert(0, '/mnt/skills/public/xlsx/scripts')
try:
    from office.soffice import get_soffice_env
except Exception:  # 没有 skills 目录时退回最简环境
    def get_soffice_env():
        env = os.environ.copy(); env['SAL_USE_VCLPLUGIN'] = 'svp'; return env

MACRO_TMPL = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Global Doc As Object
Function Sh(n As String) As Object
  Sh = Doc.Sheets.getByName(n)
End Function
Sub RunIt()
  Doc = ThisComponent
{body}
  Doc.calculateAll()
  Dim args(0) As New com.sun.star.beans.PropertyValue
  args(0).Name = "FilterName"
  args(0).Value = "Calc MS Excel 2007 XML"
  Doc.storeToURL("{out_url}", args())
  Doc.close(True)
End Sub
</script:module>"""


def _xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def run_macro(src, dst, body, timeout=1500):
    src = os.path.abspath(src); dst = os.path.abspath(dst)
    work = tempfile.mkdtemp(prefix='lo_macro_')
    try:
        prof = Path(work) / 'profile'
        env = get_soffice_env()
        subprocess.run(['soffice', '--headless', '--terminate_after_init',
                        f'-env:UserInstallation={prof.as_uri()}'],
                       capture_output=True, timeout=120, env=env)
        mdir = prof / 'user' / 'basic' / 'Standard'
        if not mdir.exists():
            raise SystemExit('LibreOffice 没建出 profile')
        # 在临时目录里放一份输入的副本，原文件不动
        tmp_in = os.path.join(work, 'in' + os.path.splitext(src)[1])
        shutil.copy(src, tmp_in)
        if os.path.exists(dst):
            os.remove(dst)
        (mdir / 'Module1.xba').write_text(
            MACRO_TMPL.format(body=_xml_escape(body), out_url=Path(dst).as_uri()), encoding='utf-8')
        t0 = time.time()
        r = subprocess.run(['timeout', str(timeout), 'soffice', '--headless', '--norestore',
                            f'-env:UserInstallation={prof.as_uri()}',
                            'vnd.sun.star.script:Standard.Module1.RunIt?language=Basic&location=application',
                            tmp_in], capture_output=True, text=True, env=env)
        if not os.path.exists(dst):
            raise SystemExit(f'宏没跑成（返回码 {r.returncode}）：{r.stderr[-800:]}')
        print(f'   宏跑完，用时 {time.time() - t0:.0f}s → {dst}')
        return dst
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print(__doc__); raise SystemExit(2)
    body = open(sys.argv[3], encoding='utf-8').read()
    run_macro(sys.argv[1], sys.argv[2], body, int(sys.argv[4]) if len(sys.argv) > 4 else 1500)
