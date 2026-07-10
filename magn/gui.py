import subprocess
import sys
import threading
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk


BASE_DIR = Path(__file__).resolve().parent
GIF_PATH = BASE_DIR / "necoarc.gif"
DONE_GIF_PATH = BASE_DIR / "necoarc2.gif"
GIF_DISPLAY_SCALE = 0.85  # GIF fills 85% of window dimension
SCRIPTS = {
    "mtmh": BASE_DIR / "mtmh",
    "keff": BASE_DIR / "keff",
    "ku": BASE_DIR / "ku",
}


class MtmhGui(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("MT/MH Processor")
        self.geometry("800x800")
        self.resizable(True, True)

        self.project_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar(value="Book1")
        self.keff_target_var = tk.StringVar(value="3")
        self.ms_target_var = tk.StringVar(value="1.5")
        self.truncate_target_var = tk.StringVar(value="2")
        self.density_var = tk.StringVar(value="7.35")
        self.ip_target_nd_var = tk.StringVar()
        self.op_target_nd_var = tk.StringVar()
        self.ip_real_nd_var = tk.StringVar()
        self.op_real_nd_var = tk.StringVar()
        self.buttons = []
        self.main_frame = None
        self.run_frame = None
        self.run_gif_label = None
        self.run_gif_frames = []
        self.done_gif_frames = []
        self.gif_running = False
        self.gif_after_id = None

        self.build_ui()

    def build_ui(self):
        self.main_frame = ttk.Frame(self, padding=14)
        self.main_frame.pack(fill="both", expand=True)

        content = ttk.Frame(self.main_frame)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=0)
        content.rowconfigure(0, weight=1)
        content.rowconfigure(1, weight=0)

        left_panel = ttk.Frame(content)
        left_panel.grid(row=0, column=0, sticky="nsew")

        right_panel = ttk.Frame(content)
        right_panel.grid(row=0, column=1, sticky="e", padx=(14, 0))
        ttk.Button(right_panel, text="逻辑", command=self.show_logic).pack(fill="x", pady=(0, 8))
        ttk.Button(right_panel, text="使用说明", command=self.show_usage).pack(fill="x", pady=(0, 8))
        ttk.Button(right_panel, text="打开文件", command=self.open_file).pack(fill="x")

        ttk.Label(content, text="群山行云\n20260531", justify="right").grid(
            row=1, column=1, sticky="se", padx=(14, 0)
        )

        file_row = ttk.Frame(left_panel)
        file_row.pack(fill="x")
        ttk.Label(file_row, text="Origin 文件").pack(side="left")
        ttk.Entry(file_row, textvariable=self.project_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(file_row, text="选择", command=self.choose_file).pack(side="left")

        sheet_row = ttk.Frame(left_panel)
        sheet_row.pack(fill="x", pady=(8, 0))
        ttk.Label(sheet_row, text="源工作表名").pack(side="left")
        ttk.Entry(sheet_row, textvariable=self.source_sheet_var, width=20).pack(side="left", padx=8)

        ttk.Label(left_panel, text="将文件中待处理工作簿预处理成：\nABC列：面内温度（K）磁场（oe）磁矩（emu/g）\nDEF列：面外温度（K）磁场（oe）磁矩（emu/g）", anchor="w", justify="left").pack(fill="x", pady=(8, 14))

        params = ttk.Frame(left_panel)
        params.pack(fill="x")
        self.add_entry(params, 0, "keff target_value", self.keff_target_var)
        self.add_entry(params, 1, "MS target x", self.ms_target_var)
        ttk.Label(params, text="不考虑退磁因子，以下不用填，也不要运行ku", anchor="w").grid(row=2, column=0, columnspan=2, sticky="w", pady=4)
        self.add_entry(params, 3, "truncate_target", self.truncate_target_var)
        self.add_entry(params, 4, "密度 ρ (g/cm³)", self.density_var)
        self.add_entry(params, 5, "ku ip_target_Nd", self.ip_target_nd_var)
        self.add_entry(params, 6, "ku op_target_Nd", self.op_target_nd_var)
        self.add_entry(params, 7, "ku ip_real_Nd", self.ip_real_nd_var)
        self.add_entry(params, 8, "ku op_real_Nd", self.op_real_nd_var)


        button_row = ttk.Frame(left_panel)
        button_row.pack(fill="x", pady=12)
        self.add_button(button_row, "运行 mtmh", lambda: self.start_run("mtmh"))
        self.add_button(button_row, "运行 keff", lambda: self.start_run("keff"))
        self.add_button(button_row, "运行 ku", lambda: self.start_run("ku"))

    def load_gif(self, path):
        """Load all GIF frames at original resolution (no subsampling)."""
        if not path.exists():
            return []

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
        """Scale a PhotoImage to fit within target dimensions.

        Uses integer zoom (enlarge) or subsample (shrink) so the result
        never exceeds target_w × target_h.
        """
        fw = max(frame.width(), 1)
        fh = max(frame.height(), 1)

        if fw <= target_w and fh <= target_h:
            # Image fits — enlarge with zoom
            zoom_x = target_w // fw
            zoom_y = target_h // fh
            factor = min(zoom_x, zoom_y)  # min keeps both dimensions within target
            if factor > 1:
                return frame.zoom(factor, factor)
            return frame
        else:
            # Image too large — shrink with subsample
            sub_x = max(1, (fw + target_w - 1) // target_w)  # ceil division
            sub_y = max(1, (fh + target_h - 1) // target_h)
            factor = max(sub_x, sub_y)  # max ensures both dimensions fit
            if factor > 1:
                return frame.subsample(factor, factor)
            return frame

    def start_gif(self):
        raw_frames = self.load_gif(GIF_PATH)
        if not raw_frames:
            return

        self.stop_gif()
        if self.main_frame is not None:
            self.main_frame.pack_forget()

        self.run_frame = ttk.Frame(self, padding=14)
        self.run_frame.pack(fill="both", expand=True)
        self.run_frame.columnconfigure(0, weight=1)
        self.run_frame.rowconfigure(0, weight=1)

        # Scale frames to fill the window
        self.update_idletasks()
        win_w = max(self.winfo_width(), 1)
        win_h = max(self.winfo_height(), 1)
        target_w = max(100, int(win_w * GIF_DISPLAY_SCALE))
        target_h = max(100, int(win_h * GIF_DISPLAY_SCALE))
        self.run_gif_frames = [self._scale_frame_to(f, target_w, target_h) for f in raw_frames]

        self.run_gif_label = ttk.Label(self.run_frame, image=self.run_gif_frames[0])
        self.run_gif_label.grid(row=0, column=0)  # centered in the cell by default

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

    def play_done_gif_once(self):
        self.stop_gif()
        self.done_gif_frames = self.load_gif(DONE_GIF_PATH)
        if not self.done_gif_frames:
            self.restore_main_frame()
            return

        # Scale frames to fill the window
        self.update_idletasks()
        win_w = max(self.winfo_width(), 1)
        win_h = max(self.winfo_height(), 1)
        target_w = max(100, int(win_w * GIF_DISPLAY_SCALE))
        target_h = max(100, int(win_h * GIF_DISPLAY_SCALE))
        self.done_gif_frames = [self._scale_frame_to(f, target_w, target_h) for f in self.done_gif_frames]

        if self.run_gif_label is not None:
            self.run_gif_label.configure(image=self.done_gif_frames[0])
        self.animate_done_gif(0)

    def animate_done_gif(self, index):
        if (
            self.run_frame is None
            or not self.run_frame.winfo_exists()
            or self.run_gif_label is None
            or not self.done_gif_frames
        ):
            return
        if index >= len(self.done_gif_frames):
            self.restore_main_frame()
            return
        self.run_gif_label.configure(image=self.done_gif_frames[index])
        self.gif_after_id = self.after(80, self.animate_done_gif, index + 1)

    def add_entry(self, parent, row, label, variable):
        ttk.Label(parent, text=label, anchor="w").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=variable, width=24).grid(row=row, column=1, sticky="w", pady=4)
        parent.grid_columnconfigure(2, weight=1)

    def add_button(self, parent, text, command):
        button = ttk.Button(parent, text=text, command=command)
        button.pack(side="left", padx=(0, 8))
        self.buttons.append(button)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Origin 项目文件",
            filetypes=[("Origin Project", "*.opju *.opj"), ("All files", "*.*")],
        )
        if path:
            self.project_var.set(path)

    def open_file(self):
        project = self.project_var.get().strip()
        if not project:
            messagebox.showwarning("提示", "请先选择 Origin 项目文件。")
            return
        path = Path(project)
        if not path.exists():
            messagebox.showerror("文件不存在", f"找不到文件：{project}")
            return
        import os
        os.startfile(path)

    def show_logic(self):
        messagebox.showinfo(
            "逻辑",
            "mtmh：读取 ipop 表，先按磁场是否连续稳定区分 MT 和 MH。MT 部分按温度最高点分成升温和降温；MH 部分按温度变化分段，少于 10 个点的段不会生成曲线。\n\n"
            "keff：读取 MHipop 表中的 H/M 曲线，按目标磁场截断左右两侧曲线，在这里截断 MH 后，会将面外 MH 归一化到面内 MH，也就是 op 乘一个系数，使得 ipop 的 mh 最后一个点对齐。然后计算 MS 和左右面积差，结果写入 MSKU。\n\n"
            "ku：读取 MHipop 表中的 H/M 曲线，先根据输入的 Nd 参数修正 H 轴，再按目标磁场截断、校正右侧曲线并计算左右面积差，结果写入 kutcxz 和 MSKU。"
        )

    def show_usage(self):
        messagebox.showinfo(
            "使用说明",


            "将待处理数据的origin文件新建表格ipop预处理成：ABC列面内温度（K）磁场（oe）磁矩（emu/g）"
            "DEF列面外温度（K）磁场（oe）磁矩（emu/g）\n\n"
            "keff target_value 为 用于算Keff的MH 截断点,算keff从0积分到这个数值\n"
            "MS target x 就是选择多大磁场作为饱和磁场，用这个点获得饱和磁矩\n\n"
            "如果不计算退磁和 ku，下面的就不要填，按顺序点 mtmh 和 keff 就行\n\n"
            "truncate target是ku的截断点，原理和keff target_value一样\n"
            "ku ip_target_Nd 和 ku op_target_Nd 是目标形状的退磁因子\n"
            "ku ip_real_Nd 和 ku op_real_Nd 是样品真实的退磁因子\n"
            "比如把一个长条修成立方体，立方体退磁因子三个面都是 0.3333，条状比如 ip 面是 0.12，op 面是 0.48\n"
            "那么就会将形状退磁因子修成一样的，方便计算 ku\n"
        )

    def start_run(self, name):
        try:
            command = self.build_command(name)
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        for button in self.buttons:
            button.config(state="disabled")
        self.start_gif()
        threading.Thread(target=self.run_named_command, args=(name, command), daemon=True).start()

    def build_command(self, name):
        project = self.validate_project()

        if name == "mtmh":
            return [
                sys.executable,
                str(SCRIPTS["mtmh"]),
                "--project",
                project,
                "--source-sheet",
                self.source_sheet_var.get().strip() or "Book1",
            ]

        if name == "keff":
            return [
                sys.executable,
                str(SCRIPTS["keff"]),
                "--project",
                project,
                "--target-value",
                self.valid_number(self.keff_target_var.get(), "keff target_value"),
                "--ms-target-x",
                self.valid_number(self.ms_target_var.get(), "MS target x"),
            ]

        if name == "ku":
            command = [
                sys.executable,
                str(SCRIPTS["ku"]),
                "--project",
                project,
                "--truncate-target",
                self.valid_number(self.truncate_target_var.get(), "truncate_target"),
                "--density",
                self.valid_number(self.density_var.get(), "密度 ρ"),
            ]
            self.add_optional_ku_args(command)
            return command

        raise ValueError(f"未知任务: {name}")

    def validate_project(self):
        project = self.project_var.get().strip()
        if not project:
            raise ValueError("请选择 Origin 项目文件。")
        if not Path(project).exists():
            raise ValueError(f"文件不存在: {project}")
        return project

    def add_optional_ku_args(self, command):
        for flag, variable, label in (
            ("--ip-target-nd", self.ip_target_nd_var, "ip_target_Nd"),
            ("--op-target-nd", self.op_target_nd_var, "op_target_Nd"),
            ("--ip-real-nd", self.ip_real_nd_var, "ip_real_Nd"),
            ("--op-real-nd", self.op_real_nd_var, "op_real_Nd"),
        ):
            value = variable.get().strip()
            if value:
                command.extend([flag, self.valid_number(value, label)])

    def valid_number(self, value, label):
        try:
            float(value)
        except ValueError as exc:
            raise ValueError(f"{label} 必须是数字。") from exc
        return value.strip()

    def run_named_command(self, name, command):
        succeeded = False
        try:
            self.run_command(command)
            succeeded = True
        except Exception as exc:
            self.after(0, self.show_run_error, exc)
        finally:
            self.after(0, self.enable_buttons, succeeded)

    def show_run_error(self, exc):
        self.restore_main_frame()
        messagebox.showerror("运行失败", self.short_error_message(exc))

    def short_error_message(self, exc):
        text = str(exc).strip()
        if not text:
            return "程序运行失败，但没有返回具体错误。"
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            return "程序运行失败，但没有返回具体错误。"

        for line in reversed(lines):
            match = re.match(r"^(?:[A-Za-z_][\w.]*Error|Exception):\s*(.+)$", line)
            if match:
                return match.group(1)
        return lines[-1]

    def enable_buttons(self, succeeded):
        if succeeded:
            self.play_done_gif_once()
        for button in self.buttons:
            button.config(state="normal")

    def run_command(self, command):
        process = subprocess.Popen(
            command,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        assert process.stdout is not None
        output = process.stdout.read()
        code = process.wait()
        if code != 0:
            message = output.strip() or f"命令退出码: {code}"
            raise RuntimeError(message)


if __name__ == "__main__":
    MtmhGui().mainloop()
