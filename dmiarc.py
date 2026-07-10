# -*- coding: utf-8 -*-
"""
Stable DMI Calculator GUI (no pandas, stable row ops)
"""

import tkinter as tk
from tkinter import filedialog, messagebox

import numpy as np
from openpyxl import Workbook
from tksheet import Sheet


# =========================================================
# 常数
# =========================================================
R = 2.612
G = 2
KB = 1.380649e-23
UB = 9.274e-24
MU0 = 4 * np.pi * 1e-7


COLUMNS = [
    "M(A/m)",
    "K(J/m3)",
    "M0(A/m)",
    "厚度t(nm)",
    "磁畴周期d(nm)",
    "拟合MT获得的B",
    "A0(J/m)",
    "Aex(J/m)",
    "DMI(mJ/m2)",
]


# =========================================================
# DMI计算
# =========================================================
def calc_dmi(M, K, M0, t_nm, d_nm, b):

    t = t_nm * 1e-9
    d = d_nm * 1e-9

    D0 = KB * (((R * G * UB) / (b * M0)) ** 0.6666667) / (4 * np.pi)
    A0 = M0 * D0 / (2 * G * UB)
    Aex = A0 * (M / M0) ** 2

    dw = 0.0
    for n in [1, 3, 5, 7, 9, 11]:
        x = 2 * np.pi * n * t / d
        w = (
            ((d ** 2) / (t ** 2))
            * (MU0 * t * (M ** 2))
            * (1 - (1 - x) * np.exp(-x))
            / ((np.pi * n) ** 3)
        )
        dw += w

    DMI = (4 * np.sqrt(K * Aex) - dw) / np.pi * 1000

    return A0, Aex, DMI


# =========================================================
# 强制结束编辑
# =========================================================
def finish_edit():
    try:
        sheet.close_text_editor(set_data=True)
    except:
        pass


# =========================================================
# 获取数据（核心稳定点）
# =========================================================
def get_data():
    finish_edit()
    return sheet.get_sheet_data()


# =========================================================
# 新增行（✔稳定版）
# =========================================================
def add_row():
    data = get_data()
    data.append([""] * len(COLUMNS))
    sheet.set_sheet_data(data)


# =========================================================
# 删除行（✔稳定版）
# =========================================================
def delete_row():

    finish_edit()

    selected = sheet.get_selected_rows()

    if not selected:
        messagebox.showinfo("提示", "请先选择行")
        return

    data = get_data()

    new_data = [
        row for i, row in enumerate(data)
        if i not in selected
    ]

    sheet.set_sheet_data(new_data)


# =========================================================
# 清空结果
# =========================================================
def clear_results():

    data = get_data()

    for i in range(len(data)):
        if len(data[i]) < 9:
            continue
        data[i][6] = ""
        data[i][7] = ""
        data[i][8] = ""

    sheet.set_sheet_data(data)


# =========================================================
# 计算
# =========================================================
def run_calculation():

    data = get_data()

    error_rows = []

    for i, row in enumerate(data):

        try:
            if all(str(x).strip() == "" for x in row[:6]):
                continue

            M = float(row[0])
            K = float(row[1])
            M0 = float(row[2])
            t_nm = float(row[3])
            d_nm = float(row[4])
            b = float(row[5])

            A0, Aex, DMI = calc_dmi(M, K, M0, t_nm, d_nm, b)

            data[i][6] = f"{A0:.6e}"
            data[i][7] = f"{Aex:.6e}"
            data[i][8] = f"{DMI:.6f}"

        except:
            error_rows.append(i + 1)

    sheet.set_sheet_data(data)

    if error_rows:
        messagebox.showwarning("部分失败", f"错误行：{error_rows}")
    else:
        messagebox.showinfo("完成", "计算完成")


# =========================================================
# 导出 Excel（无 pandas）
# =========================================================
def export_excel():

    data = get_data()

    clean = [
        row for row in data
        if any(str(x).strip() != "" for x in row)
    ]

    if not clean:
        messagebox.showinfo("提示", "没有数据")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")]
    )

    if not path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "DMI"

    ws.append(COLUMNS)

    for row in clean:
        ws.append(row)

    wb.save(path)

    messagebox.showinfo("成功", "导出完成")


# =========================================================
# GUI
# =========================================================
root = tk.Tk()
root.title("DMI Calculator 群山行云")
root.geometry("1350x750")


# ---------- 按钮 ----------
top = tk.Frame(root)
top.pack(fill="x", pady=5)

tk.Button(top, text="新增行", width=12, command=add_row).pack(side="left", padx=5)
tk.Button(top, text="删除行", width=12, command=delete_row).pack(side="left", padx=5)
tk.Button(top, text="计算", width=12, command=run_calculation).pack(side="left", padx=5)
tk.Button(top, text="清空结果", width=12, command=clear_results).pack(side="left", padx=5)
tk.Button(top, text="导出Excel", width=12, command=export_excel).pack(side="left", padx=5)


# ---------- 表格 ----------
sheet = Sheet(
    root,
    headers=COLUMNS,
    data=[[""] * len(COLUMNS) for _ in range(20)]
)

sheet.enable_bindings((
    "single_select",
    "row_select",
    "column_select",
    "arrowkeys",
    "edit_cell",
    "copy",
    "paste",
    "delete",
))

sheet.pack(fill="both", expand=True)


# ---------- 底部文本 ----------
bottom = tk.Frame(root)
bottom.pack(fill="x", pady=8)

text_box = tk.Text(bottom, height=5, font=("Arial", 12))
text_box.pack(fill="x", padx=10)

text_box.insert(
    "1.0",
    "DMI Calculator\n"
    "Zhang et al.Above-room-temperature chiral skyrmion lattice and Dzyaloshinskii–Moriya interaction in a van der Waals ferromagnet Fe3−xGaTe2, Nat. Commun. 2024\n"
    "畴壁能：Electric-field-driven non-volatile multi-state switching of individual skyrmions in a multiferroic heterostructure\n"
    "群山行云 2026.5.13"
)


# =========================================================
# 启动
# =========================================================
root.mainloop()