"""共享组件：GridTable, 滚动容器, 智能搜索, 全局滚轮"""
import tkinter as tk
from tkinter import ttk


class GridTable:

    """基于 Entry 网格的表格，支持单元格独立着色"""

    def __init__(self, parent, headers, rows, cell_styles=None, edit_mode=False):

        self.parent = parent

        self.headers = list(headers)

        self.rows = [list(r) for r in rows]

        self.edit_mode = edit_mode

        self.cell_styles = cell_styles or {}

        self.sel_cells = set()

        self.anchor = None

        self.ncols = len(headers)

        self.nrows = len(rows)

        self._drag_moved = False


        self.frame = tk.Frame(parent, bg="#d0d0d0")

        self._build()


    def _build(self):

        for w in self.frame.winfo_children():

            w.destroy()

        ncols = max(self.ncols, 1)

        nrows = self.nrows

        col_w = max(80, min(180, 900 // ncols))


        # 表头

        for ci in range(ncols):

            h = self.headers[ci] if ci < len(self.headers) else f"列{ci + 1}"

            if self.edit_mode:

                w = tk.Entry(self.frame, font=("微软雅黑", 11, "bold"),

                             justify="center", relief="flat", bg="#d9e1e8")

                w.insert(0, h)

            else:

                w = tk.Label(self.frame, text=h, font=("微软雅黑", 11, "bold"),

                             bg="#d9e1e8", fg="#2c3e50", anchor="center")

            w.grid(row=0, column=ci, sticky="nsew", padx=1, pady=1)

            self.frame.columnconfigure(ci, weight=1, minsize=col_w)


        # 数据行

        for ri in range(nrows):

            for ci in range(ncols):

                val = self.rows[ri][ci] if ci < len(self.rows[ri]) else ""

                style = self.cell_styles.get((ri, ci), {})

                if self.edit_mode:

                    w = tk.Entry(self.frame, font=("微软雅黑", 11),

                                 relief="flat", justify="left",

                                 bg=style.get("bg", "white"),

                                 fg=style.get("fg", "black"))

                    w.insert(0, str(val))

                    w.bind("<ButtonPress-1>", lambda e, r=ri, c=ci: self._on_press(r, c, e))

                    w.bind("<B1-Motion>", lambda e, r=ri, c=ci: self._on_move(r, c))

                    w.bind("<Button-3>", lambda e, r=ri, c=ci: self._on_right_click(r, c, e))

                    w.bind("<Tab>", lambda e, r=ri, c=ci: self._on_tab(r, c))

                    w.bind("<Shift-Tab>", lambda e, r=ri, c=ci: self._on_shift_tab(r, c))

                else:

                    w = tk.Label(self.frame, text=str(val), font=("微软雅黑", 11),

                                 bg=style.get("bg", "white"),

                                 fg=style.get("fg", "#2c3e50"),

                                 anchor="w", padx=4)

                w.grid(row=ri + 1, column=ci, sticky="nsew", padx=1, pady=1)


    def _apply_selection(self):

        for ri in range(self.nrows):

            for ci in range(self.ncols):

                style = self.cell_styles.get((ri, ci), {})

                bg = "#cce5ff" if (ri, ci) in self.sel_cells else style.get("bg", "white")

                fg = style.get("fg", "black")

                for w in self.frame.grid_slaves(row=ri + 1, column=ci):

                    try: w.configure(bg=bg, fg=fg)

                    except Exception: pass


    def _on_press(self, ri, ci, event):

        self._drag_moved = False

        if event.state & 0x0004:

            if (ri, ci) in self.sel_cells:

                self.sel_cells.discard((ri, ci))

            else:

                self.sel_cells.add((ri, ci))

        elif event.state & 0x20000 and self.anchor:

            ar, ac = self.anchor

            lo_r, hi_r = min(ar, ri), max(ar, ri)

            lo_c, hi_c = min(ac, ci), max(ac, ci)

            self.sel_cells.clear()

            for r in range(lo_r, hi_r + 1):

                for c in range(lo_c, hi_c + 1):

                    self.sel_cells.add((r, c))

        else:

            self.sel_cells.clear()

            self.sel_cells.add((ri, ci))

        self.anchor = (ri, ci)

        self._apply_selection()


    def _on_move(self, ri, ci):

        if self.anchor and (ri, ci) != self.anchor:

            self._drag_moved = True

            ar, ac = self.anchor

            lo_r, hi_r = min(ar, ri), max(ar, ri)

            lo_c, hi_c = min(ac, ci), max(ac, ci)

            self.sel_cells.clear()

            for r in range(lo_r, hi_r + 1):

                for c in range(lo_c, hi_c + 1):

                    self.sel_cells.add((r, c))

            self._apply_selection()


    def _on_tab(self, ri, ci):

        nc = ci + 1; nr = ri

        if nc >= self.ncols: nc = 0; nr = ri + 1

        if nr >= self.nrows: nr = 0

        self.sel_cells.clear(); self.sel_cells.add((nr, nc))

        self.anchor = (nr, nc); self._apply_selection()

        for w in self.frame.grid_slaves(row=nr + 1, column=nc):

            w.focus_set(); return "break"

        return "break"


    def _on_shift_tab(self, ri, ci):

        nc = ci - 1; nr = ri

        if nc < 0: nc = self.ncols - 1; nr = ri - 1

        if nr < 0: nr = self.nrows - 1

        self.sel_cells.clear(); self.sel_cells.add((nr, nc))

        self.anchor = (nr, nc); self._apply_selection()

        for w in self.frame.grid_slaves(row=nr + 1, column=nc):

            w.focus_set(); return "break"

        return "break"


    def _on_right_click(self, ri, ci, event):

        if (ri, ci) not in self.sel_cells:

            self.sel_cells.clear(); self.sel_cells.add((ri, ci))

            self.anchor = (ri, ci); self._apply_selection()

        menu = tk.Menu(self.frame, tearoff=0)

        count = len(self.sel_cells)

        label = f"（{count}格）" if count > 1 else ""

        fmt_menu = tk.Menu(menu, tearoff=0)

        def _bg():

            from tkinter import colorchooser

            c = colorchooser.askcolor(title=f"选择{count}个格子的背景色")

            if c and c[1]:

                for (r, c2) in list(self.sel_cells):

                    if (r, c2) not in self.cell_styles: self.cell_styles[(r, c2)] = {}

                    self.cell_styles[(r, c2)]["bg"] = c[1]

                self._apply_selection()

        fmt_menu.add_command(label=f"填充背景色...{label}", command=_bg)

        def _fg():

            from tkinter import colorchooser

            c = colorchooser.askcolor(title=f"选择{count}个格子的文字色")

            if c and c[1]:

                for (r, c2) in list(self.sel_cells):

                    if (r, c2) not in self.cell_styles: self.cell_styles[(r, c2)] = {}

                    self.cell_styles[(r, c2)]["fg"] = c[1]

                self._apply_selection()

        fmt_menu.add_command(label=f"文字颜色...{label}", command=_fg)

        fmt_menu.add_separator()

        if len(self.sel_cells) >= 2:

            fmt_menu.add_command(label="🔗 合并选中格子", command=self._merge_cells)

        menu.add_cascade(label="单元格格式", menu=fmt_menu)

        menu.post(event.x_root, event.y_root)


    def _sync_entries(self):

        """将 Entry 中的当前值写回 self.rows"""

        if not self.edit_mode:

            return

        for ri in range(self.nrows):

            for ci in range(self.ncols):

                for w in self.frame.grid_slaves(row=ri + 1, column=ci):

                    if isinstance(w, tk.Entry):

                        val = w.get().strip()

                        while len(self.rows[ri]) <= ci:

                            self.rows[ri].append("")

                        self.rows[ri][ci] = val


    def get_data(self):

        self._sync_entries()

        hdrs = []

        for ci in range(self.ncols):

            txt = self.headers[ci] if ci < len(self.headers) else f"列{ci + 1}"

            if self.edit_mode:

                for w in self.frame.grid_slaves(row=0, column=ci):

                    if isinstance(w, tk.Entry): txt = w.get().strip() or txt

            hdrs.append(txt)

        data_rows = []

        for ri in range(self.nrows):

            row_vals = list(self.rows[ri])

            while len(row_vals) < self.ncols: row_vals.append("")

            data_rows.append(row_vals[:self.ncols])

        styles_export = {}

        for (r, c), s in self.cell_styles.items():

            styles_export[f"{r},{c}"] = s

        return hdrs, data_rows, styles_export


    def _rebuild(self):

        self.ncols = len(self.headers)

        self.nrows = len(self.rows)

        for w in self.frame.winfo_children(): w.destroy()

        self._build()

        self._apply_selection()


    def add_row_at_end(self):

        self._sync_entries()

        self.rows.append([""] * self.ncols); self._rebuild()


    def insert_row_above(self):

        self._sync_entries()

        if self.sel_cells:

            r = min(rc[0] for rc in self.sel_cells)

            self.rows.insert(r, [""] * self.ncols)

        else:

            self.rows.append([""] * self.ncols)

        self._rebuild()


    def delete_selected_rows(self):

        self._sync_entries()

        rows_to_del = set(r for (r, c) in self.sel_cells)

        if not rows_to_del: return

        for r in sorted(rows_to_del, reverse=True):

            if 0 <= r < len(self.rows): self.rows.pop(r)

        self.sel_cells.clear(); self._rebuild()


    def add_column_at_end(self):

        self._sync_entries()

        self.headers.append(f"列{self.ncols + 1}")

        for row in self.rows:

            while len(row) < len(self.headers): row.append("")

        self._rebuild()


    def delete_last_column(self):

        if self.ncols <= 1: return

        self._sync_entries()

        self.headers.pop()

        for row in self.rows:

            if len(row) > len(self.headers): row.pop()

        self._rebuild()


    def _merge_cells(self):

        if len(self.sel_cells) < 2: return

        self._sync_entries()

        sc = sorted(self.sel_cells)

        fr, fc = sc[0]

        # 确保 rows 有足够列

        while len(self.rows[fr]) <= fc:

            self.rows[fr].append("")

        combined = self.rows[fr][fc]

        for r, c in sc[1:]:

            while len(self.rows) <= r:

                self.rows.append([""] * self.ncols)

            while len(self.rows[r]) <= c:

                self.rows[r].append("")

            if self.rows[r][c].strip():

                combined += str(self.rows[r][c])

            self.rows[r][c] = ""

        self.rows[fr][fc] = combined

        self.sel_cells.clear()

        self.sel_cells.add((fr, fc))

        self._rebuild()

