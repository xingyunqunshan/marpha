"""
统一启动面板 — 9 个按钮调用所有工具
开发环境直接跑 python，打包后用 bundled 方式运行子脚本
"""
import os
import subprocess
import sys
from pathlib import Path

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# =========================================================
# 路径解析
# =========================================================
if getattr(sys, "frozen", False):
    # PyInstaller 打包后
    _MEIPASS = Path(sys._MEIPASS)
    BASE = _MEIPASS
    # bundled python: 取 pythonw.exe（无控制台窗口）
    _PYTHON = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(_PYTHON):
        _PYTHON = os.path.join(_MEIPASS, "pythonw.exe")
    if not os.path.exists(_PYTHON):
        # fallback: 把当前 exe 当 python 用（PyInstaller exe 支持 argv[1] 为脚本名）
        _PYTHON = sys.executable
    _FROZEN = True
else:
    BASE = Path(__file__).resolve().parent
    _PYTHON = sys.executable
    _FROZEN = False


PROGRAMS = [
    # (显示名称, 相对路径, 说明)
    ("十七分割",   "hallsep/gui.py",           "单温度Hall效应贡献分离"),
    ("无限剑制",    "magn/gui.py",              "MH/MT温度分割+参数计算"),
    ("矛盾螺旋",  "sym/gui.py",               "Hall/MR温度分割+对称化处理"),
    ("HEAVEN'S FEEL",        "LTEM/LTEM.py",     "LTEM 数据处理"),
    ("根源",     "dmiarc.py",                "DMI计算"),
    ("星之内海",   "edx.py",                   "EDX数据转PPT"),
    ("开辟之星",   "omftoppt_mpsarrowfix.py",  "omf,ovf 转 PPT"),
    ("痛觉残留",    "ppt123tobigppt.py",        "多 PPT 合并"),
    ("空之境界",  "topocharge.py",            "ovf计算拓扑荷"),
]


def _make_runner_script(script_abs: str, script_dir: str) -> str:
    """生成一个 runner .py 文件，内容是把 sys.path 配好后 exec 目标脚本"""
    runner = os.path.join(script_dir, ".__avalon_runner.py")
    with open(runner, "w", encoding="utf-8") as f:
        f.write(f'import sys, os\n')
        f.write(f'sys.path[0:0] = {repr([script_dir])}\n')
        f.write(f'os.chdir({repr(script_dir)})\n')
        if _FROZEN:
            # frozen 模式: 把 MEIPASS 的 _internal 也加进去
            f.write(f'sys.path.insert(0, {repr(str(_MEIPASS))})\n')
            blz = os.path.join(str(_MEIPASS), "base_library.zip")
            if os.path.exists(blz):
                f.write(f'sys.path.insert(0, {repr(blz)})\n')
        f.write(f'exec(open({repr(script_abs)}, encoding="utf-8").read())\n')
    return runner


def run_program(relative_path):
    """在子进程中启动指定脚本"""
    script_abs = str(BASE / relative_path)
    script_dir = str(BASE / os.path.dirname(relative_path))

    runner = _make_runner_script(script_abs, script_dir)

    def cleanup(r=runner):
        try:
            os.remove(r)
        except OSError:
            pass

    # subprocess 跑完后清理 runner
    proc = subprocess.Popen(
        [_PYTHON, "-u", runner],
        cwd=script_dir,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    # 用线程等进程结束后删除 runner
    def _wait_and_cleanup():
        try:
            proc.wait(timeout=600)
        except subprocess.TimeoutExpired:
            proc.kill()
        finally:
            cleanup()

    import threading
    threading.Thread(target=_wait_and_cleanup, daemon=True).start()


# =========================================================
# GUI
# =========================================================
root = ttk.Window(
    title="Marble Phantasm",
    themename="cosmo",
    size=(900, 1000),
    resizable=(False, False)
)

# 标题
title_label = ttk.Label(
    root,
    text="Marble Phantasm",
    font=("Arial", 20, "bold"),
    anchor=CENTER
)
title_label.pack(pady=(24, 6))

subtitle = ttk.Label(
    root,
    text="空想具现化",
    font=("Arial", 11),
    anchor=CENTER,
    bootstyle=SECONDARY
)
subtitle.pack(pady=(0, 16))

# 按钮区域
btn_frame = ttk.Frame(root, padding=20)
btn_frame.pack(fill=BOTH, expand=YES)

# 3 列布局
COLS = 3
ROWS = 3

for idx, (name, path, desc) in enumerate(PROGRAMS):
    row = idx // COLS
    col = idx % COLS

    card = ttk.Labelframe(
        btn_frame,
        text=name,
        padding=12,
        bootstyle=INFO
    )
    card.pack_propagate(False)   # 固定大小，不随内容收缩
    card.grid(
        row=row,
        column=col,
        padx=10,
        pady=10,
        ipadx=4,
        ipady=4,
        sticky="nsew"
    )

    # 说明文字（expand 撑满，把按钮推到底部）
    ttk.Label(
        card,
        text=desc,
        font=("Arial", 10),
        wraplength=140,
        anchor=CENTER,
        justify=CENTER,
        bootstyle=SECONDARY
    ).pack(fill=BOTH, expand=True, pady=(0, 10))

    # 启动按钮（底部对齐）
    btn = ttk.Button(
        card,
        text="▶  启动",
        bootstyle=SUCCESS,
        width=14,
        command=lambda p=path: run_program(p)
    )
    btn.pack(side=BOTTOM, anchor=CENTER)

# 均匀分布行列
for c in range(COLS):
    btn_frame.columnconfigure(c, weight=1, uniform="col")
for r in range(ROWS):
    btn_frame.rowconfigure(r, weight=1, uniform="row")

# 底部
footer = ttk.Label(
    root,
    text="     群山行云 · 2026\n基于版本python 3.14.3",
    font=("Arial", 10),
    bootstyle=SECONDARY,
    anchor=CENTER
)
footer.pack(pady=(4, 12))

# =========================================================
# 运行
# =========================================================
root.mainloop()
