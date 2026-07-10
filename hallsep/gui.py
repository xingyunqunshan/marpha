"""可视化 GUI — ttkbootstrap 美化 + GIF 动效 + subprocess 运行"""
import os
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk

BASE_DIR = Path(__file__).resolve().parent
GIF_PATH = BASE_DIR / "necoarc3.gif"
SCRIPT_PATH = BASE_DIR / "2.py"
GIF_DISPLAY_SCALE = 0.85  # GIF 占窗口比例（参照 mtmh）


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("自动霍尔效应分离")
        self.geometry("820x820")
        self.resizable(True, True)

        self.project_var = tk.StringVar()
        self.tol_var = tk.StringVar(value="100")
        self.fill_e_var = tk.BooleanVar(value=True)
        self.source_var = tk.StringVar(value="Book1")

        self.buttons = []
        self.main_frame = None
        self.run_frame = None
        self.run_gif_label = None
        self.run_gif_frames = []
        self.gif_running = False
        self.gif_after_id = None

        # 预先加载 GIF 帧，避免点击后卡顿
        self._preloaded_frames = self.load_gif(GIF_PATH)

        self.build_ui()

    # ────────── UI ──────────

    def build_ui(self):
        self.main_frame = ttk.Frame(self, padding=16)
        self.main_frame.pack(fill="both", expand=True)

        # 上半部分：参数区（占约1/3高度）
        top_frame = ttk.Frame(self.main_frame)
        top_frame.pack(fill="x", pady=(0, 0))

        # ── 文件选择 ──
        file_row = ttk.Frame(top_frame)
        file_row.pack(fill="x", pady=(0, 8))
        ttk.Label(file_row, text="Origin 项目文件", width=14).pack(side="left")
        ttk.Entry(file_row, textvariable=self.project_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(file_row, text="选择", command=self.choose_file).pack(side="left")

        # ── 工作表名 ──
        sheet_row = ttk.Frame(top_frame)
        sheet_row.pack(fill="x", pady=(0, 8))
        ttk.Label(sheet_row, text="源工作表名", width=14).pack(side="left")
        ttk.Entry(sheet_row, textvariable=self.source_var, width=16).pack(side="left", padx=8)
        ttk.Label(sheet_row, text="→ 输出表自动加 after 后缀", font=("", 9)).pack(side="left")

        # ── 容差参数 ──
        tol_row = ttk.Frame(top_frame)
        tol_row.pack(fill="x", pady=(0, 8))
        ttk.Label(tol_row, text="MH和ρxyH匹配容错（oe）", width=22).pack(side="left")
        ttk.Entry(tol_row, textvariable=self.tol_var, width=12).pack(side="left", padx=8)
        ttk.Checkbutton(tol_row, text="E列ρxx填充1（不影响结果）", variable=self.fill_e_var).pack(side="left", padx=(12, 0))

        # ── 分隔 ──
        ttk.Separator(top_frame).pack(fill="x", pady=(12, 8))

        # ── 操作按钮行 ──
        btn_row = ttk.Frame(top_frame)
        btn_row.pack(fill="x", pady=(0, 4))

        self.add_button(btn_row, "① 复制 → 转换", lambda: self.start_run("copy"), bootstyle="primary")
        self.add_button(btn_row, "② 匹配 → 拟合", lambda: self.start_run("match"), bootstyle="primary")
        self.add_button(btn_row, "③ 画图", lambda: self.start_run("plot"), bootstyle="primary")
        self.add_button(btn_row, "▶ 一键全跑", lambda: self.start_run("all"), bootstyle="success")

        # ── 信息按钮行（下移，单独一行） ──
        info_row = ttk.Frame(top_frame)
        info_row.pack(fill="x", pady=(4, 0))
        self.add_button(info_row, "逻辑", self.show_logic, bootstyle="info-outline")
        self.add_button(info_row, "使用说明", self.show_usage, bootstyle="info-outline")
        self.add_button(info_row, "打开选定文件", self.open_selected_file, bootstyle="info-outline")

    def show_logic(self):
        messagebox.showinfo(
            "处理逻辑",
            "━━━ ① 复制 → 转换 ━━━\n"
            "从源工作表复制全部列到 源表名+after。\n"
            "• A、C列 ÷ 10000（OE→T 单位转换）\n"
            "• E列：勾选时填充1，不勾选时保留原始值\n"
            "• F列 = E × 1\n"
            "• G列 = F × F\n\n"
            "━━━ ② 匹配 → 拟合 ━━━\n"
            "匹配 A列 与 C列（容差内视为匹配）：\n"
            "• A、C列各自找到最小值（谷底）作为分界点\n"
            "• 前段（谷底之前）：A前段 × C前段 匹配\n"
            "• 后段（谷底之后）：A后段 × C后段 匹配\n"
            "→ 输出 J～AA 列（J=ρxyH, K=ρxx, L=MH_H, M=MH_M,\n"
            "   N=MH_ρxy, Q=N×K/J, R=M/J）\n"
            "• 去重：J列、L列各自保留 |J-L| 最小的行\n"
            "• 对 Q×R 前14点线性回归，得斜率(U4)和截距(U5)\n"
            "→ 计算 W～AA 衍生列：\n"
            "   W=L, X=M, Y=J×U5, Z=Q×U4×J, AA=X-Y-Z\n\n"
            "━━━ ③ 画图 ━━━\n"
            "以 W(B) 为X轴，X(rouxy)/Y(NHE)/Z(AHE)/AA(THE) 为Y轴\n"
            "生成点线图，自动配色，图例标签来自列注释。\n"
            "图名 = 源表名graph"
        )

    def show_usage(self):
        """使用说明，用户可自行修改文本内容。"""
        messagebox.showinfo(
            "使用说明",
           
            "0. 适用于PPMS的单温度hall曲线贡献分离，需要事先分割温度\n"
            "1. 准备好 Origin 项目文件（.opju）\n"
            "2. AB(M-H曲线)，A：H(oe)，B：M(emu/g)\n"
            "3. CDE(ρ-H曲线)，C：H(oe)，D：ρxy(μΩcm),E：ρxxρ(μΩcm)\n"
            "4. 匹配容错指的是ρ-H，M-H使用的测试磁场并不总相同，挑选出同磁场下的数据ABCEG放在JKLMN\n"
            "5. E列可以不填，使用自动填充，不影响分离结果，看公式就知道了\n"
            "6. 操作完之后自动保存并关闭 Origin\n"
            "7. 测试5个样本正常工作"
        )

    def open_selected_file(self):
        """用系统默认程序打开已选定的 Origin 项目文件。"""
        p = self.project_var.get().strip()
        if not p:
            messagebox.showwarning("未选文件", "请先在上方选择 Origin 项目文件。")
            return
        path = Path(p)
        if not path.exists():
            messagebox.showerror("文件不存在", f"找不到文件：{p}")
            return
        try:
            os.startfile(str(path.resolve()))
        except Exception as exc:
            messagebox.showerror("打开失败", f"无法打开文件：{exc}")

    def add_button(self, parent, text, command, bootstyle="secondary"):
        btn = ttk.Button(parent, text=text, command=command, bootstyle=bootstyle)
        btn.pack(side="left", padx=(0, 8))
        self.buttons.append(btn)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Origin 项目文件",
            filetypes=[("Origin Project", "*.opju *.opj"), ("All Files", "*.*")],
        )
        if path:
            self.project_var.set(path)

    # ────────── GIF 动效 ──────────

    def load_gif(self, path):
        """加载 GIF 的所有帧（原始尺寸）。"""
        frames = []
        index = 0
        while True:
            try:
                frame = tk.PhotoImage(file=str(path), format=f"gif -index {index}")
            except tk.TclError:
                break
            frames.append(frame)
            index += 1
        return frames

    def _scale_frame_to(self, frame, target_w, target_h):
        """把帧缩放到目标尺寸（整数 zoom/subsample）。"""
        fw = max(frame.width(), 1)
        fh = max(frame.height(), 1)
        if fw <= target_w and fh <= target_h:
            zoom_x = target_w // fw
            zoom_y = target_h // fh
            factor = min(zoom_x, zoom_y)
            if factor > 1:
                return frame.zoom(factor, factor)
            return frame
        else:
            sub_x = max(1, (fw + target_w - 1) // target_w)
            sub_y = max(1, (fh + target_h - 1) // target_h)
            factor = max(sub_x, sub_y)
            if factor > 1:
                return frame.subsample(factor, factor)
            return frame

    def start_gif(self):
        """显示 GIF 动效，缩放帧填满窗口。"""
        raw_frames = self._preloaded_frames or self.load_gif(GIF_PATH)
        if not raw_frames:
            return
        self.stop_gif()
        if self.main_frame is not None:
            self.main_frame.pack_forget()

        self.run_frame = ttk.Frame(self, padding=14)
        self.run_frame.pack(fill="both", expand=True)
        self.run_frame.columnconfigure(0, weight=1)
        self.run_frame.rowconfigure(0, weight=1)

        # 缩放帧（只做一次，不在动画循环里做）
        self.update_idletasks()
        win_w = max(self.winfo_width(), 1)
        win_h = max(self.winfo_height(), 1)
        target_w = max(100, int(win_w * GIF_DISPLAY_SCALE))
        target_h = max(100, int(win_h * GIF_DISPLAY_SCALE))
        self.run_gif_frames = [self._scale_frame_to(f, target_w, target_h) for f in raw_frames]

        self.run_gif_label = ttk.Label(self.run_frame, image=self.run_gif_frames[0])
        self.run_gif_label.grid(row=0, column=0)

        self.gif_running = True
        self.animate_loop_gif(0)

    def stop_gif(self):
        self.gif_running = False
        if self.gif_after_id is not None:
            self.after_cancel(self.gif_after_id)
            self.gif_after_id = None

    def restore_main_frame(self):
        self.stop_gif()
        if self.run_frame is not None and self.run_frame.winfo_exists():
            self.run_frame.destroy()
        self.run_frame = None
        self.run_gif_label = None
        if self.main_frame is not None and not self.main_frame.winfo_ismapped():
            self.main_frame.pack(fill="both", expand=True)

    def animate_loop_gif(self, index):
        if not self.gif_running or not self.run_gif_frames or self.run_gif_label is None:
            return
        self.run_gif_label.configure(image=self.run_gif_frames[index])
        next_index = (index + 1) % len(self.run_gif_frames)
        self.gif_after_id = self.after(80, self.animate_loop_gif, next_index)

    def play_done_once(self):
        """运行成功：停 GIF，保持最后一帧 1.2 秒后恢复。"""
        self.stop_gif()
        self.after(1200, self.restore_main_frame)

    # ────────── 运行控制 ──────────

    def get_tol(self):
        try:
            val = float(self.tol_var.get().strip())
            if val <= 0:
                raise ValueError
            return val
        except ValueError:
            messagebox.showerror("参数错误", "容差必须是正数。")
            raise

    def validate_project(self):
        p = self.project_var.get().strip()
        if not p:
            messagebox.showerror("错误", "请先选择 Origin 项目文件。")
            raise ValueError("no file")
        path = Path(p)
        if not path.exists():
            messagebox.showerror("错误", f"文件不存在: {p}")
            raise ValueError("no file")
        return path.resolve()

    def set_buttons_state(self, enabled):
        state = "normal" if enabled else "disabled"
        for btn in self.buttons:
            btn.config(state=state)

    def start_run(self, step):
        """构建命令，启动 GIF，在后台线程跑 subprocess。"""
        try:
            project = str(self.validate_project())
            tol = self.get_tol()
        except (ValueError, SystemExit):
            return

        source = self.source_var.get().strip() or "Book1"
        target = source + "after"

        command = [
            sys.executable, str(SCRIPT_PATH),
            "--step", step,
            "--tol", str(tol),
            "--source-book", source,
            "--target-book", target,
        ]
        if not self.fill_e_var.get():
            command.append("--no-fill-e")
        command.append(project)

        self.set_buttons_state(False)
        self.start_gif()
        threading.Thread(target=self._run_command, args=(command,), daemon=True).start()

    def _run_command(self, command):
        """在后台线程中跑 subprocess（参照 mtmh/gui.py）。"""
        succeeded = False
        try:
            proc = subprocess.Popen(
                command,
                cwd=str(BASE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert proc.stdout is not None
            output = proc.stdout.read()
            code = proc.wait()
            if code != 0:
                msg = output.strip() or f"命令退出码: {code}"
                raise RuntimeError(msg)
            succeeded = True
        except Exception as exc:
            self.after(0, self._show_error, exc)
        finally:
            self.after(0, self._enable_after_run, succeeded)

    def _show_error(self, exc):
        self.restore_main_frame()
        messagebox.showerror("运行失败", str(exc))

    def _enable_after_run(self, succeeded):
        if succeeded:
            self.play_done_once()
        else:
            self.restore_main_frame()
        self.set_buttons_state(True)

    def run(self):
        self.mainloop()


if __name__ == "__main__":
    App().run()
