# -*- coding: utf-8 -*-
import os
import re
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk


BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "Tspilt"
GIF_PATH = BASE_DIR / "necoarc3.gif"
GIF_DISPLAY_SCALE = 0.85  # GIF fills 85% of window dimension


class SymGui(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Hall/MR 对称化处理")
        self.geometry("720x700")
        self.resizable(True, True)

        self.project_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar(value="Book1")
        self.max_field_var = tk.StringVar(value="30000")
        self.pair_tolerance_var = tk.StringVar(value="200")
        self.interpolate_var = tk.BooleanVar(value=True)
        self.buttons = []
        self.main_frame = None
        self.run_frame = None
        self.run_gif_label = None
        self.run_gif_frames = []
        self.gif_running = False
        self.gif_after_id = None

        self.build_ui()

    def build_ui(self):
        self.main_frame = ttk.Frame(self, padding=14)
        self.main_frame.pack(fill="both", expand=True)

        file_row = ttk.Frame(self.main_frame)
        file_row.pack(fill="x")
        ttk.Label(file_row, text="Origin 文件").pack(side="left")
        ttk.Entry(file_row, textvariable=self.project_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(file_row, text="选择", command=self.choose_file).pack(side="left")

        params = ttk.Frame(self.main_frame)
        params.pack(fill="x", pady=(14, 8))
        self.add_entry(params, 0, "源工作表", self.source_sheet_var)
        self.add_entry(params, 1, "最大磁场 Oe", self.max_field_var)
        self.add_entry(params, 2, "配对容差 Oe", self.pair_tolerance_var)
        ttk.Checkbutton(
            params,
            text="插值：每两个相邻点之间插入平均磁场和平均电阻",
            variable=self.interpolate_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=8)

        # Two-column workflow layout
        workflow_row = ttk.Frame(self.main_frame)
        workflow_row.pack(fill="x", pady=(10, 6))
        workflow_row.columnconfigure(0, weight=1)
        workflow_row.columnconfigure(1, weight=1)

        hall_frame = ttk.LabelFrame(workflow_row, text="Hall 处理")
        hall_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        hall_inner = ttk.Frame(hall_frame, padding=8)
        hall_inner.pack(fill="both", expand=True)
        self.add_button(hall_inner, "生成 sym / sym-chazhi", self.start_generate)
        self.add_button(hall_inner, "sym-fix 配对修正", self.start_fix)
        self.add_button(hall_inner, "Hall 对称化", lambda: self.start_simple("--hall-sym", "霍尔对称化完成"))
        self.add_button(hall_inner, "Hall 符号反转", lambda: self.start_simple("--reverse-hall", "霍尔符号反转完成"))

        mr_frame = ttk.LabelFrame(workflow_row, text="MR 处理")
        mr_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        mr_inner = ttk.Frame(mr_frame, padding=8)
        mr_inner.pack(fill="both", expand=True)
        self.add_button(mr_inner, "生成 sym / sym-chazhi", self.start_generate)
        self.add_button(mr_inner, "sym-fix 配对修正", self.start_fix)
        self.add_button(mr_inner, "MR 归一化", lambda: self.start_simple("--mr-normalize", "MR 归一化完成"))

        ttk.Label(
            self.main_frame,
            text=(
                "输入表前三列应为：温度、磁场、电阻。\n"
                "勾选插值时，后续按钮处理 sym-chazhi；不勾选时，后续按钮处理 sym。\n"
                "每列从上到下依次执行：生成 → 配对修正 → 对称化/归一化。"
            ),
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(10, 0))

        info_row = ttk.Frame(self.main_frame)
        info_row.pack(fill="x", pady=(10, 0))
        ttk.Button(info_row, text="逻辑", command=self.show_logic).pack(side="left")
        ttk.Button(info_row, text="打开文件", command=self.open_selected_file).pack(side="left", padx=(8, 0))

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

    def add_entry(self, parent, row, label, variable):
        ttk.Label(parent, text=label, anchor="w").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=variable, width=28).grid(row=row, column=1, sticky="w", pady=5)
        parent.grid_columnconfigure(2, weight=1)

    def add_button(self, parent, text, command):
        button = ttk.Button(parent, text=text, command=command)
        button.pack(fill="x", pady=2)
        self.buttons.append(button)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Origin 项目文件",
            filetypes=[("Origin Project", "*.opju *.opj"), ("All files", "*.*")],
        )
        if path:
            self.project_var.set(path)

    def show_logic(self):
        messagebox.showinfo(
            "处理逻辑",
            "注意，只适用于PPMS测出来的没有多余数据的变温变场Hall/MR数据，如用源表数据中保温的需要手动去掉\n\n"
            
            "第一步 — 生成 sym / sym-chazhi：\n"
            "读取输入表中温度、磁场、电阻三列数据，按温度分组。每组内以零磁场为界，将数据拆分为正场（H≥0）和负场（H≤0）两支。\n"
            "若勾选插值，则在相邻两个数据点之间插入二者的平均磁场和平均电阻，生成 sym-chazhi 表；若不勾选，则直接生成 sym 表。\n\n"
            "第二步 — sym-fix 配对修正：\n"
            "对上一步生成的 sym 或 sym-chazhi 表进行配对检查。确保正场和负场两侧的数据点一一对应（磁场绝对值相近），\n"
            "对无法配对或配对异常的组进行标记和修正，修正结果写入 fix 表。\n\n"
            "第三步（Hall）— Hall 对称化：\n"
            "读取 fix 表中的 Hall 数据，对正负磁场下的电阻值进行反对称处理：\n"
            "R_sym(H) = [R(H) − R(−H)] / 2，消除纵向电阻分量的混入，得到纯净的霍尔电阻。\n\n"
            "第三步（Hall 可选）— Hall 符号反转：\n"
            "将霍尔电阻数据整体取反（R → −R），用于修正测量时电极接线方向与约定相反的情况。\n\n"
            "第三步（MR）— MR 归一化：\n"
            "读取 fix 表中的 MR 数据，将每个温度组内所有电阻值除以零磁场处的电阻 R(0)，\n"
            "得到归一化磁电阻 MR(%) = [R(H) − R(0)] / R(0) × 100%。\n"
            "共测试4组数据，2026.7.8结束。"
        )

    def open_selected_file(self):
        """用系统默认程序打开已选定的 Origin 文件。"""
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

    def start_generate(self):
        try:
            command = self.build_generate_command()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.start_command(command, "生成完成")

    def start_fix(self):
        try:
            source_sheet = self.validate_source_sheet()
            command = [
                sys.executable,
                str(SCRIPT_PATH),
                "--project",
                self.validate_project(),
                "--source-sheet",
                source_sheet,
                "--fix-pairs",
                "--pair-tolerance",
                self.valid_number(self.pair_tolerance_var.get(), "配对容差 Oe"),
            ]
            if self.interpolate_var.get():
                command.append("--interpolate")
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.start_command(command, "sym-fix 完成")

    def start_simple(self, flag, success_title):
        try:
            command = [
                sys.executable,
                str(SCRIPT_PATH),
                "--project",
                self.validate_project(),
                "--source-sheet",
                self.validate_source_sheet(),
                flag,
            ]
            if self.interpolate_var.get():
                command.append("--interpolate")
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.start_command(command, success_title)

    def start_command(self, command, success_title):
        for button in self.buttons:
            button.config(state="disabled")
        self.start_gif()
        threading.Thread(target=self.run_thread, args=(command, success_title), daemon=True).start()

    def validate_project(self):
        project = self.project_var.get().strip()
        if not project:
            raise ValueError("请选择 Origin 项目文件。")
        if not Path(project).exists():
            raise ValueError(f"文件不存在: {project}")
        return project

    def validate_source_sheet(self):
        source_sheet = self.source_sheet_var.get().strip()
        if not source_sheet:
            raise ValueError("源工作表不能为空。")
        return source_sheet

    def build_generate_command(self):
        source_sheet = self.validate_source_sheet()

        command = [
            sys.executable,
            str(SCRIPT_PATH),
            "--project",
            self.validate_project(),
            "--source-sheet",
            source_sheet,
            "--max-field",
            self.valid_number(self.max_field_var.get(), "最大磁场 Oe"),
        ]
        if self.interpolate_var.get():
            command.append("--interpolate")
        return command

    def valid_number(self, value, label):
        try:
            float(value)
        except ValueError as exc:
            raise ValueError(f"{label} 必须是数字。") from exc
        return value.strip()

    def run_thread(self, command, success_title):
        try:
            output = self.run_command(command)
        except Exception as exc:
            self.after(0, self.show_error, exc)
        else:
            self.after(0, self.show_success, success_title, output)
        finally:
            self.after(0, self.enable_buttons)

    def run_command(self, command):
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        process = subprocess.Popen(
            command,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        assert process.stdout is not None
        output = process.stdout.read()
        code = process.wait()
        if code != 0:
            raise RuntimeError(output.strip() or f"命令退出码: {code}")
        return output

    def show_success(self, title, output):
        self.restore_main_frame()
        messagebox.showinfo(title, self.success_summary(output))

    def show_error(self, exc):
        self.restore_main_frame()
        messagebox.showerror("运行失败", self.short_error_message(exc))

    def enable_buttons(self):
        for button in self.buttons:
            button.config(state="normal")

    def success_summary(self, text):
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        translations = {
            "Groups processed:": "已处理组数:",
            "Interrupted groups:": "中断组数:",
            "Hall symmetry groups:": "霍尔对称组数:",
            "MR normalized groups:": "MR 归一化组数:",
            "MR F-filled groups:": "MR F 填充组数:",
            "MR skipped groups:": "MR 跳过组数:",
            "MR plots:": "MR 图表数:",
            "Hall sign reversed groups:": "霍尔符号反转组数:",
            "Hall plots:": "霍尔图表数:",
        }
        summary = []
        for line in lines:
            if line.startswith("Interrupted group indexes:"):
                group_text = line[len("Interrupted group indexes:"):].strip()
                if group_text:
                    groups = [item.strip() for item in group_text.split(",") if item.strip()]
                    summary.append(f"第{'、'.join(groups)}组中断，手动检查")
                continue
            for en_prefix, zh_prefix in translations.items():
                if line.startswith(en_prefix):
                    summary.append(zh_prefix + line[len(en_prefix):])
                    break
        return "\n".join(summary) if summary else "处理完成。"

    def short_error_message(self, exc):
        text = str(exc).strip()
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            return "程序运行失败，但没有返回具体错误。"

        for line in reversed(lines):
            match = re.match(r"^(?:[A-Za-z_][\w.]*Error|Exception):\s*(.+)$", line)
            if match:
                return match.group(1)
        return lines[-1]


if __name__ == "__main__":
    SymGui().mainloop()
