import os
import numpy as np
import matplotlib.pyplot as plt
import discretisedfield as df

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog

# =========================================================
# 主窗口 — 浅色主题
# =========================================================
root = ttk.Window(
    title="Topological Density GUI",
    themename="cosmo",          # 浅色主题
    size=(1000, 300),
    resizable=(False, False)
)

folder_var = ttk.StringVar()

# =========================================================
# 选择文件夹
# =========================================================
def select_folder():

    folder = filedialog.askdirectory(
        title="选择OMF/OVF文件夹"
    )

    if folder:
        folder_var.set(folder)

# =========================================================
# 主程序
# =========================================================
def calculate():

    input_folder = folder_var.get()

    if not input_folder:
        Messagebox.show_error("请先选择文件夹", "Error")
        return

    # =====================================================
    # 输出文件夹
    # =====================================================
    output_folder = os.path.join(input_folder, "output")

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # 搜索文件
    # =====================================================
    files = []

    for f in os.listdir(input_folder):

        lower = f.lower()

        if lower.endswith(".omf") or lower.endswith(".ovf"):
            files.append(f)

    files.sort()

    if not files:
        Messagebox.show_error("未找到OMF/OVF文件", "Error")
        return

    print(f"Found {len(files)} files")

    # =====================================================
    # 主循环
    # =====================================================
    for file in files:

        print("\n=================================================")
        print("Processing:", file)

        input_file = os.path.join(input_folder, file)

        # =================================================
        # 读取
        # =================================================
        field = df.Field.from_file(input_file)

        data = field.array

        Nx, Ny, Nz, _ = data.shape

        print("Data shape =", data.shape)

        # =================================================
        # 中层
        # =================================================
        z_index = Nz // 2

        print("Using z index =", z_index)

        m = data[:, :, z_index, :]

        mx = m[:, :, 0]
        my = m[:, :, 1]
        mz = m[:, :, 2]

        # =================================================
        # 归一化
        # =================================================
        norm = np.sqrt(mx**2 + my**2 + mz**2)

        norm[norm == 0] = 1

        mx = mx / norm
        my = my / norm
        mz = mz / norm

        # =================================================
        # 拓扑密度
        # =================================================
        dmx_dx = np.gradient(mx, axis=0)
        dmy_dx = np.gradient(my, axis=0)
        dmz_dx = np.gradient(mz, axis=0)

        dmx_dy = np.gradient(mx, axis=1)
        dmy_dy = np.gradient(my, axis=1)
        dmz_dy = np.gradient(mz, axis=1)

        q = (
            mx * (dmy_dx * dmz_dy - dmz_dx * dmy_dy)
            + my * (dmz_dx * dmx_dy - dmx_dx * dmz_dy)
            + mz * (dmx_dx * dmy_dy - dmy_dx * dmx_dy)
        )

        Q = np.sum(q) / (4 * np.pi)

        print(f"Topological charge Q = {Q:.6f}")

        # =================================================
        # 文件名（带Q）
        # =================================================
        name = os.path.splitext(file)[0]

        q_text = f"{Q:.3f}"

        output_png = os.path.join(
            output_folder,
            f"{name}_Q_{q_text}.png"
        )

        output_ovf = os.path.join(
            output_folder,
            f"{name}_Q_{q_text}.ovf"
        )

        # =================================================
        # 图像变换
        #
        # 左右镜像
        # + 顺时针90°
        # =================================================
        q_show = np.fliplr(q)
        q_show = np.rot90(q_show, k=1)

        mx_show = np.fliplr(mx)
        mx_show = np.rot90(mx_show, k=1)

        my_show = np.fliplr(my)
        my_show = np.rot90(my_show, k=1)

        # =================================================
        # PNG
        # =================================================
        Nx2, Ny2 = q_show.shape

        stride = max(Nx2 // 40, 1)

        X, Y = np.meshgrid(
            np.arange(Ny2),
            np.arange(Nx2)
        )

        plt.figure(figsize=(8, 8), dpi=300)

        # =================================================
        # 背景
        # =================================================
        plt.imshow(
            q_show,
            cmap="seismic",
            origin="lower"
        )

        # =================================================
        # 箭头
        # =================================================
        plt.quiver(
            X[::stride, ::stride],
            Y[::stride, ::stride],

            mx_show[::stride, ::stride],
            my_show[::stride, ::stride],

            color="k",

            pivot="mid",

            scale=40,
            width=0.003,

            headwidth=4,
            headlength=5,
            headaxislength=5
        )

        plt.title(f"{name}   Q={Q:.4f}")

        plt.colorbar(label="q(x,y)")

        plt.tight_layout()

        plt.savefig(
            output_png,
            bbox_inches="tight"
        )

        plt.close()

        print("PNG saved:")
        print(output_png)

        # =================================================
        # 单层OVF
        # =================================================
        new_data = np.zeros((Nx, Ny, 1, 3))

        new_data[:, :, 0, 0] = mx
        new_data[:, :, 0, 1] = my
        new_data[:, :, 0, 2] = mz

        # =================================================
        # mesh
        # =================================================
        cell = field.mesh.cell

        pmin = field.mesh.region.pmin
        pmax = field.mesh.region.pmax

        new_mesh = df.Mesh(
            region=df.Region(
                p1=(pmin[0], pmin[1], 0),
                p2=(pmax[0], pmax[1], cell[2])
            ),
            cell=(cell[0], cell[1], cell[2])
        )

        # =================================================
        # field
        # =================================================
        new_field = df.Field(
            mesh=new_mesh,
            nvdim=3,
            value=new_data
        )

        # =================================================
        # 保存OVF
        # =================================================
        new_field.to_file(output_ovf)

        print("OVF saved:")
        print(output_ovf)

    Messagebox.show_info("全部处理完成", "Done")

# =========================================================
# About 弹窗
# =========================================================
def show_about():
    Messagebox.show_info(
        "Q为拓扑荷，输出图的箭头方向为磁矩方向，大小表示拓扑密度。\n\n群山行云\n20260517",
        "关于"
    )

# =========================================================
# GUI布局
# =========================================================
frame = ttk.Frame(root)
frame.pack(pady=20)

btn_select = ttk.Button(
    frame,
    text="选择文件夹",
    bootstyle=OUTLINE,
    width=24,
    command=select_folder
)
btn_select.pack(pady=6)

entry = ttk.Entry(
    frame,
    textvariable=folder_var,
    width=55,
    font=("Arial", 12)
)
entry.pack(pady=6)

btn_calc = ttk.Button(
    frame,
    text="计算",
    bootstyle=SUCCESS,
    width=24,
    command=calculate
)
btn_calc.pack(pady=12)

# =========================================================
# About 按钮
# =========================================================
btn_about = ttk.Button(
    root,
    text="关于",
    bootstyle=SECONDARY,
    width=12,
    command=show_about
)

btn_about.pack(
    side="bottom",
    pady=12
)

# =========================================================
# 运行
# =========================================================
root.mainloop()