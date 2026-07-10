from pathlib import Path
import argparse
import math
import sys

import originpro as op


SOURCE_BOOK = "Book1"
TARGET_BOOK = "Book1after"
TOL_BEFORE = 100
TOL_AFTER = 100
FILL_E = True
OUTPUT_HEADERS = ["h", "m", "h", "rouxy", "rouxx2", "N*K/J", "M/J",
                "B", "rouxy", "NHE", "AHE", "THE"]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Copy Book1 to Book1after in an OPJU file and write matched A/B/C/D/G data to J:N."
    )
    parser.add_argument("project", nargs="?", help="Path to .opju/.opj file.")
    parser.add_argument("--source-book", default=SOURCE_BOOK)
    parser.add_argument("--target-book", default=TARGET_BOOK)
    parser.add_argument("--tol", type=float, default=None, help="匹配容差 (同时设置 TOL_BEFORE 和 TOL_AFTER)")
    parser.add_argument("--step", choices=["copy", "match", "plot", "all"], default="all",
                        help="执行步骤: copy=复制转换, match=匹配拟合, plot=画图, all=全流程")
    parser.add_argument("--fill-e", action="store_true", default=True, help="E列填充1 (默认开启)")
    parser.add_argument("--no-fill-e", action="store_false", dest="fill_e", help="不填充E列")
    return parser.parse_args()


def select_project_file():
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        path = Path(sys.argv[1]).expanduser()
        if path.exists():
            return path.resolve()

    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="选择 Origin 项目文件",
        filetypes=[("Origin Project", "*.opju *.opj"), ("All Files", "*.*")],
        parent=root,
    )
    root.destroy()

    if not path:
        raise SystemExit("未选择文件。")
    return Path(path).resolve()


def iter_workbooks():
    index = 0
    while True:
        book = op.find_book("w", index)
        if book is None:
            break
        yield book
        index += 1


def find_book(name):
    for book in iter_workbooks():
        names = {str(getattr(book, "name", "")), str(getattr(book, "lname", ""))}
        if name in names:
            return book
    return None


def get_book_or_raise(name):
    book = find_book(name)
    if book is None:
        available = ", ".join(
            f"{book.name}" + (f" ({book.lname})" if book.lname and book.lname != book.name else "")
            for book in iter_workbooks()
        )
        raise Exception(f"项目中找不到工作簿 {name}。当前工作簿: {available or '无'}")
    return book


def get_first_sheet(book):
    try:
        return book[0]
    except Exception as exc:
        raise Exception(f"工作簿 {book.name} 中没有可用工作表。") from exc


def recreate_target_book(name):
    old = find_book(name)
    if old is not None:
        old.destroy()

    wks = op.new_sheet("w", lname=name)
    book = wks.get_book()
    book.name = name
    book.lname = name
    return book, wks


def safe_shape(wks):
    rows, cols = wks.shape
    return int(rows or 0), int(cols or 0)


def copy_label(src, dst, src_col, dst_col):
    for kind in ("L", "U", "C"):
        try:
            value = src.get_label(src_col, kind)
        except Exception:
            value = ""
        if value:
            try:
                dst.set_label(dst_col, value, kind)
            except Exception:
                pass


def copy_sheet(src, dst):
    _, col_count = safe_shape(src)
    if col_count < 1:
        raise Exception("Book1 没有可复制的数据列。")

    dst._check_add_cols(col_count)
    for col in range(col_count):
        values = list(src.to_list(col))
        dst.from_list(col, values)
        copy_label(src, dst, col, col)


def to_float(value):
    if value is None:
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def read_numeric_column(wks, col):
    return [to_float(value) for value in wks.to_list(col)]


def trim_to_data_rows(columns):
    last = -1
    for values in columns:
        for index, value in enumerate(values):
            if value is not None:
                last = max(last, index)
    if last < 1:
        return []

    rows = []
    for index in range(1, last + 1):
        rows.append([values[index] if index < len(values) else None for values in columns])
    return rows


def find_valley(rows, col_index):
    """找到数据列的最小值索引（V形谷底），作为前/后段自动分界点。"""
    best_idx = 0
    best_val = None
    for index, row in enumerate(rows):
        value = row[col_index]
        if value is not None:
            if best_val is None or value < best_val:
                best_val = value
                best_idx = index
    return best_idx


def build_matches(rows):
    if not rows:
        return []

    # 自动找到 A列 和 C列 各自的谷底，作为前/后段分界
    x_valley = find_valley(rows, 0)
    y_valley = find_valley(rows, 2)
    # 谷底归入前段（下降段），后段从谷底之后开始（上升段）
    x_end = x_valley + 1
    y_end = y_valley + 1
    print(f"A列谷底: rows[{x_valley}]={rows[x_valley][0]:.4f}, 分界x_end={x_end}")
    print(f"C列谷底: rows[{y_valley}]={rows[y_valley][2]:.4f}, 分界y_end={y_end}")

    matches = []

    # 前段匹配：A前段 × C前段
    for i in range(x_end):
        a = rows[i][0]
        if a is None:
            continue
        for j in range(y_end):
            c = rows[j][2]
            if c is None:
                continue
            if abs(a - c) < TOL_BEFORE:
                k, m, n = rows[i][1], rows[j][3], rows[j][6]
                q = n * k / a if a and k is not None and n is not None else None
                r = m / a if a and m is not None else None
                matches.append([a, k, c, m, n, q, r])

    # 后段匹配：A后段 × C后段
    for i in range(x_end, len(rows)):
        a = rows[i][0]
        if a is None:
            continue
        for j in range(y_end, len(rows)):
            c = rows[j][2]
            if c is None:
                continue
            if abs(a - c) < TOL_AFTER:
                k, m, n = rows[i][1], rows[j][3], rows[j][6]
                q = n * k / a if a and k is not None and n is not None else None
                r = m / a if a and m is not None else None
                matches.append([a, k, c, m, n, q, r])

    return matches


def dedup_matches(matches):
    """去重：J列相同值时保留|J-L|最小的行；L列相同值时保留|J-L|最小的行。"""
    if not matches:
        return matches

    # 第一步：按J列去重 —— 相同J值只保留与L最接近的那行
    best_by_j = {}
    for row in matches:
        j_val = row[0]
        diff = abs(j_val - row[2])
        if j_val not in best_by_j or diff < best_by_j[j_val][1]:
            best_by_j[j_val] = (row, diff)
    deduped = [entry[0] for entry in best_by_j.values()]

    # 第二步：按L列去重 —— 相同L值只保留与J最接近的那行
    best_by_l = {}
    for row in deduped:
        l_val = row[2]
        diff = abs(row[0] - l_val)
        if l_val not in best_by_l or diff < best_by_l[l_val][1]:
            best_by_l[l_val] = (row, diff)

    return [entry[0] for entry in best_by_l.values()]


# 列映射: J=9,K=10,L=11,M=12,N=13, O/P留空, Q=16,R=17, S留空, T=19,U=20, V留空, W=22,X=23,Y=24,Z=25,AA=26
OUTPUT_COL_MAP = [9, 10, 11, 12, 13, 16, 17, 22, 23, 24, 25, 26]
OUTPUT_MAX_COL = 26
FIT_COUNT = 14


def clear_output_columns(wks, row_count):
    clear_rows = max(row_count, len(wks.to_list(9)), 1)
    blanks = [""] * clear_rows
    for col in OUTPUT_COL_MAP:
        wks.from_list(col, blanks)


def write_matches(wks, matches, slope=None, intercept=None):
    wks._check_add_cols(OUTPUT_MAX_COL + 1)
    clear_output_columns(wks, len(matches) + 1)

    header_row = list(OUTPUT_HEADERS)
    ext_rows = []
    for row in matches:
        j_val, k_val, l_val, m_val, n_val, q_val, r_val = row
        w_val = l_val
        x_val = m_val
        if slope is not None and intercept is not None and j_val is not None and q_val is not None and k_val is not None and m_val is not None:
            y_val = j_val * intercept   # Y = J * U5 (U5=截距)
            z_val = q_val * slope * j_val  # Z = Q * U4 * J (U4=斜率)
            aa_val = x_val - y_val - z_val
        else:
            y_val, z_val, aa_val = None, None, None
        ext_rows.append([j_val, k_val, l_val, m_val, n_val, q_val, r_val,
                         w_val, x_val, y_val, z_val, aa_val])

    output = [header_row] + ext_rows
    for offset, col_idx in enumerate(OUTPUT_COL_MAP):
        wks.from_list(col_idx, [row[offset] for row in output])

    for offset, label in enumerate(OUTPUT_HEADERS):
        wks.set_label(OUTPUT_COL_MAP[offset], label, "L")

    # 为绘图列设置 Comment(C) 标签，Origin 图例会用到
    plot_comments = {22: "B", 23: "rouxy", 24: "NHE", 25: "AHE", 26: "THE"}
    for col, comment in plot_comments.items():
        wks.set_label(col, comment, "C")


def linear_fit(q_vals, r_vals):
    """对前 FIT_COUNT 个 (Q, R) 点做线性回归 R = slope * Q + intercept。"""
    x = q_vals[:FIT_COUNT]
    y = r_vals[:FIT_COUNT]
    n = len(x)

    sum_x = sum(x)
    sum_y = sum(y)
    sum_xy = sum(xi * yi for xi, yi in zip(x, y))
    sum_x2 = sum(xi * xi for xi in x)

    slope = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x * sum_x)
    intercept = (sum_y - slope * sum_x) / n

    y_mean = sum_y / n
    ss_res = sum((yi - (slope * xi + intercept)) ** 2 for xi, yi in zip(x, y))
    ss_tot = sum((yi - y_mean) ** 2 for yi in y)
    r_squared = 1 - ss_res / ss_tot if ss_tot != 0 else 0

    return slope, intercept, r_squared, n


def write_fit_results(wks, matches, slope, intercept, r_squared):
    """拟合结果写入 T、U 列。"""
    if slope is None:
        print("无拟合结果，跳过 T、U 列写入。")
        return

    n = min(FIT_COUNT, len(matches))

    t_lines = [
        f"回归方程: y = {slope:.6f}x + {intercept:.6f}",
        f"R² = {r_squared:.6f}",
        f"拟合点数: {n}",
        f"斜率: {slope:.6f}",
        f"截距: {intercept:.6f}",
    ]

    u_lines = [
        "",
        r_squared,
        n,
        slope,
        intercept,
    ]

    wks._check_add_cols(OUTPUT_MAX_COL + 1)
    t_existing = list(wks.to_list(19))
    u_existing = list(wks.to_list(20))
    max_len = max(len(t_lines), len(u_lines))

    for i in range(max_len):
        t_val = t_lines[i] if i < len(t_lines) else ""
        u_val = u_lines[i] if i < len(u_lines) else ""
        if i < len(t_existing):
            t_existing[i] = t_val
            u_existing[i] = u_val
        else:
            t_existing.append(t_val)
            u_existing.append(u_val)

    wks.from_list(19, t_existing)
    wks.from_list(20, u_existing)
    wks.set_label(19, "拟合结果", "L")
    wks.set_label(20, "拟合数值", "L")

    print(f"线性拟合 (前{n}点): y = {slope:.6f}x + {intercept:.6f}, R^2 = {r_squared:.6f}")


def format_graph_layer(gl):
    """照搬 mtmh 的图层格式化：四轴显示、边框宽度3、刻度5个。"""
    gl.lt_exec(
        "layer.x.showAxes=3;"
        "layer.y.showAxes=3;"
        "layer.x2.showAxes=3;"
        "layer.y2.showAxes=3;"
        "layer.x.ticks=5;"
        "layer.y.ticks=5;"
        "layer.x2.ticks=5;"
        "layer.y2.ticks=5;"
        "layer.x.thickness=3;"
        "layer.y.thickness=3;"
        "layer.x2.thickness=3;"
        "layer.y2.thickness=3;"
        "layer.frame.width=3;"
    )


def create_plot(wks):
    """以W列为x轴，X/Y/Z/AA为y轴作图 —— 参照 mt.py 写法。"""
    x_col = 22
    y_cols = [23, 24, 25, 26]

    gp = op.new_graph(template="Scatter")
    gp.lname = SOURCE_BOOK + "graph"
    gl = gp[0]

    plots = []
    for yc in y_cols:
        plots.append(gl.add_plot(wks, coly=yc, colx=x_col))

    for p in plots:
        p.lt = 200  # 点线图

    if len(plots) > 1:
        gl.group()  # 自动分配颜色，图例使用列 Comment 标签

    gl.rescale()
    format_graph_layer(gl)

    gl.lt_exec(
        "layer.x.label$ = 'B';"
        "layer.y.label$ = '';"
        "layer.x.ticksin=1;"
        "layer.y.ticksin=1;"
        "layer.x2.ticksin=1;"
        "layer.y2.ticksin=1;"
        "legend.show=1;"
        "legend.update();"
    )

    print("已创建图表: X轴=W(B), Y轴=rouxy/NHE/AHE/THE")


def transform_data(wks):
    """复制后对数据做单位转换和衍生列计算。"""
    # 确保至少有 7 列（A~G），否则 F/G 列无法写入
    wks._check_add_cols(7)

    # A列 /10000
    a_data = list(wks.to_list(0))
    a_trans = [v / 10000 if isinstance(v, (int, float)) else v for v in a_data]
    wks.from_list(0, a_trans)

    # C列 /10000
    c_data = list(wks.to_list(2))
    c_trans = [v / 10000 if isinstance(v, (int, float)) else v for v in c_data]
    wks.from_list(2, c_trans)

    # E列填充1 (与D列对齐，默认开启)
    if FILL_E:
        d_data = list(wks.to_list(3))
        # 找到D列最后有效行，E列填充相同行数的1
        d_len = len([v for v in d_data if v is not None and not (isinstance(v, str) and v.strip() == "")])
        e_ones = [1] * max(d_len, 1)
        wks.from_list(4, e_ones)

    # F列 = E列 * 1
    e_data = list(wks.to_list(4))
    f_data = [v * 1 if isinstance(v, (int, float)) else v for v in e_data]
    wks.from_list(5, f_data)

    # G列 = F列 * F列
    g_data = [v * v if isinstance(v, (int, float)) else v for v in f_data]
    wks.from_list(6, g_data)

    e_info = "E=1" if FILL_E else "E不变"
    print(f"已完成: A/10000, C/10000, {e_info}, F=E*1, G=F*F")


def process_project(project_path, source_book, target_book):
    op.set_show(True)
    op.open(str(project_path))

    src_book = get_book_or_raise(source_book)
    src_wks = get_first_sheet(src_book)
    _, dst_wks = recreate_target_book(target_book)
    copy_sheet(src_wks, dst_wks)
    transform_data(dst_wks)

    dst_wks._check_add_cols(OUTPUT_MAX_COL + 1)
    columns = [read_numeric_column(dst_wks, col) for col in range(7)]
    rows = trim_to_data_rows(columns)
    matches = build_matches(rows)
    before_dedup = len(matches)
    matches = dedup_matches(matches)
    print(f"去重: {before_dedup} -> {len(matches)} 条 (删除 {before_dedup - len(matches)} 条重复)")

    q_vals = [row[5] for row in matches if row[5] is not None and row[6] is not None]
    r_vals = [row[6] for row in matches if row[5] is not None and row[6] is not None]
    if len(q_vals) >= 2:
        slope, intercept, r_squared, _ = linear_fit(q_vals, r_vals)
        print(f"线性拟合 (前{FIT_COUNT}点): y = {slope:.6f}x + {intercept:.6f}, R^2 = {r_squared:.6f}")
    else:
        slope, intercept, r_squared = None, None, None
        print("拟合数据不足，跳过。")

    write_matches(dst_wks, matches, slope, intercept)
    write_fit_results(dst_wks, matches, slope, intercept, r_squared)
    create_plot(dst_wks)
    op.lt_exec("system.project.backup=0;system.project.autosaveon=0;")
    if not op.save(str(project_path)):
        raise Exception(f"保存失败，请确认文件没有被占用且有写入权限: {project_path}")

    return len(rows), len(matches)


def _apply_tol(tol):
    if tol is not None:
        global TOL_BEFORE, TOL_AFTER
        TOL_BEFORE = tol
        TOL_AFTER = tol


def _save_and_exit(project_path):
    op.lt_exec("system.project.backup=0;system.project.autosaveon=0;")
    if not op.save(str(project_path)):
        raise Exception(f"保存失败，请确认文件没有被占用且有写入权限: {project_path}")
    print("已保存。")


def step_copy(project_path):
    """仅复制+转换。"""
    op.set_show(False)
    op.open(str(project_path))
    src_book = get_book_or_raise(SOURCE_BOOK)
    src_wks = get_first_sheet(src_book)
    _, dst_wks = recreate_target_book(TARGET_BOOK)
    copy_sheet(src_wks, dst_wks)
    transform_data(dst_wks)
    _save_and_exit(project_path)
    print("复制+转换 完成。")


def step_match(project_path):
    """匹配+拟合（若 Book1after 不存在则先复制转换）。"""
    op.set_show(False)
    op.open(str(project_path))

    dst_book = find_book(TARGET_BOOK)
    if dst_book is None:
        src_book = get_book_or_raise(SOURCE_BOOK)
        src_wks = get_first_sheet(src_book)
        _, dst_wks = recreate_target_book(TARGET_BOOK)
        copy_sheet(src_wks, dst_wks)
        transform_data(dst_wks)
    else:
        dst_wks = dst_book[0]

    dst_wks._check_add_cols(OUTPUT_MAX_COL + 1)
    columns = [read_numeric_column(dst_wks, col) for col in range(7)]
    rows = trim_to_data_rows(columns)
    matches = build_matches(rows)
    before_dedup = len(matches)
    matches = dedup_matches(matches)
    print(f"去重: {before_dedup} -> {len(matches)} 条 (删除 {before_dedup - len(matches)} 条重复)")

    qv = [r[5] for r in matches if r[5] is not None and r[6] is not None]
    rv = [r[6] for r in matches if r[5] is not None and r[6] is not None]
    if len(qv) >= 2:
        slope, icpt, r2, _ = linear_fit(qv, rv)
    else:
        slope, icpt, r2 = None, None, None

    write_matches(dst_wks, matches, slope, icpt)
    write_fit_results(dst_wks, matches, slope, icpt, r2)
    _save_and_exit(project_path)
    print(f"匹配+拟合 完成 ({len(matches)} 条)。")


def step_plot(project_path):
    """仅画图。"""
    op.set_show(False)
    op.open(str(project_path))

    dst_book = find_book(TARGET_BOOK)
    if dst_book is None:
        raise Exception("Book1after 不存在，请先执行匹配。")
    dst_wks = dst_book[0]
    create_plot(dst_wks)
    _save_and_exit(project_path)
    print("画图 完成。")


def step_all(project_path):
    """全流程：复制+转换+匹配+拟合+画图。"""
    op.set_show(True)
    op.open(str(project_path))

    src_book = get_book_or_raise(SOURCE_BOOK)
    src_wks = get_first_sheet(src_book)
    _, dst_wks = recreate_target_book(TARGET_BOOK)
    copy_sheet(src_wks, dst_wks)
    transform_data(dst_wks)

    dst_wks._check_add_cols(OUTPUT_MAX_COL + 1)
    columns = [read_numeric_column(dst_wks, col) for col in range(7)]
    rows = trim_to_data_rows(columns)
    matches = build_matches(rows)
    before_dedup = len(matches)
    matches = dedup_matches(matches)
    print(f"去重: {before_dedup} -> {len(matches)} 条 (删除 {before_dedup - len(matches)} 条重复)")

    qv = [r[5] for r in matches if r[5] is not None and r[6] is not None]
    rv = [r[6] for r in matches if r[5] is not None and r[6] is not None]
    if len(qv) >= 2:
        slope, icpt, r2, _ = linear_fit(qv, rv)
        print(f"线性拟合 (前{FIT_COUNT}点): y = {slope:.6f}x + {icpt:.6f}, R^2 = {r2:.6f}")
    else:
        slope, icpt, r2 = None, None, None
        print("拟合数据不足，跳过。")

    write_matches(dst_wks, matches, slope, icpt)
    write_fit_results(dst_wks, matches, slope, icpt, r2)
    create_plot(dst_wks)
    _save_and_exit(project_path)
    return len(rows), len(matches)


def main():
    # 确保控制台输出使用 UTF-8，避免中文乱码
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    args = parse_args()
    _apply_tol(args.tol)
    global FILL_E, SOURCE_BOOK, TARGET_BOOK
    FILL_E = args.fill_e
    SOURCE_BOOK = args.source_book
    TARGET_BOOK = args.target_book

    project_path = Path(args.project).resolve() if args.project else select_project_file()
    if project_path.suffix.lower() not in (".opju", ".opj"):
        raise SystemExit("请选择 .opju 或 .opj 文件。")

    try:
        if args.step == "copy":
            step_copy(project_path)
        elif args.step == "match":
            step_match(project_path)
        elif args.step == "plot":
            step_plot(project_path)
        else:
            row_count, match_count = step_all(project_path)
            print(f"完成: 已复制 {args.source_book} 为 {args.target_book}")
            print(f"读取数据行: {row_count}")
            print(f"匹配输出: {match_count} 条，写入 {args.target_book} 的 J:AA")
    finally:
        op.exit()


if __name__ == "__main__":
    main()
