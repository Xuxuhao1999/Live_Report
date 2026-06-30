"""服务记录标签页：可编辑表格 + Excel 导出"""
import os
import json
from datetime import date, datetime
from tkinter import ttk, messagebox, filedialog
import tkinter as tk


class ServiceRecordMixin:
        def build_service_record(self, parent):
            """构建服务记录标签页"""
            # ---- 数据文件路径 ----
            self.service_data_file = os.path.join(self.app_dir, "service_records.json")
    
            # ---- 事项下拉选项 ----
            self.service_items = [
                "发布雷电黄色预警信号", "发布雷电橙色预警信号", "发布雷电红色预警信号",
                "发布暴雨黄色预警信号", "发布暴雨橙色预警信号", "发布暴雨红色预警信号",
                "发布大风黄色预警信号", "发布大风橙色预警信号", "发布大风红色预警信号",
                "发布雷暴大风黄色预警信号", "发布雷暴大风橙色预警信号", "发布雷暴大风红色预警信号",
                "入驻防指",
            ]
    
            # ---- 服务方式选项 ----
            self.service_modes = ["浙政钉", "微信群", "电话", "短信", "闪信", "语音外呼"]
    
            # ---- 头部：日期 + 天气类型 ----
            header_frame = ttk.Frame(parent, padding=10)
            header_frame.pack(fill="x")
    
            today = date.today()
            ttk.Label(header_frame, text="", font=("微软雅黑", 13, "bold")).pack(side="left")  # spacer
            self.svc_year = ttk.Combobox(header_frame, values=[str(y) for y in range(today.year - 2, today.year + 3)],
                                         width=5, state="normal")
            self.svc_year.set(str(today.year))
            self.svc_year.pack(side="left")
            ttk.Label(header_frame, text="年", font=("微软雅黑", 12)).pack(side="left")
            self.svc_month = ttk.Combobox(header_frame, values=[f"{m:02d}" for m in range(1, 13)],
                                          width=4, state="normal")
            self.svc_month.set(f"{today.month:02d}")
            self.svc_month.pack(side="left")
            ttk.Label(header_frame, text="月", font=("微软雅黑", 12)).pack(side="left")
            self.svc_day = ttk.Combobox(header_frame, values=[f"{d:02d}" for d in range(1, 32)],
                                        width=4, state="normal")
            self.svc_day.set(f"{today.day:02d}")
            self.svc_day.pack(side="left")
            ttk.Label(header_frame, text="日", font=("微软雅黑", 12)).pack(side="left")
    
            self.svc_weather = ttk.Combobox(header_frame, values=["强对流", "台风", "暴雨", "大风"],
                                            width=6, state="normal")
            self.svc_weather.set("")
            self.svc_weather.pack(side="left", padx=5)
            ttk.Label(header_frame, text="天气服务记录", font=("微软雅黑", 13, "bold")).pack(side="left", padx=5)
    
            # ---- 表格 ----
            table_frame = ttk.Frame(parent, padding=5)
            table_frame.pack(fill="both", expand=True)
    
            columns = ("序号", "时间", "事项", "内容", "服务方式", "备注")
            # 增大行高便于阅读
            svc_style = ttk.Style()
            svc_style.configure("Svc.Treeview", font=("微软雅黑", 11), rowheight=36)
            self.svc_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                         selectmode="extended", height=12, style="Svc.Treeview")
            col_widths = [50, 120, 180, 300, 160, 120]
            for col, w in zip(columns, col_widths):
                self.svc_tree.heading(col, text=col)
                self.svc_tree.column(col, width=w, anchor="center" if col in ("序号", "时间") else "w", minwidth=40)
    
            # 滚动条
            vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.svc_tree.yview)
            hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.svc_tree.xview)
            self.svc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            self.svc_tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            table_frame.rowconfigure(0, weight=1)
            table_frame.columnconfigure(0, weight=1)
    
            # 单击编辑（改为 ButtonRelease 以兼容拖动选择）
            self.svc_tree.bind("<ButtonRelease-1>", self._on_svc_click_or_edit)
    
            # 双击表头编辑
            self.svc_tree.bind("<Double-1>", self._on_svc_double_click)
    
            # 右键菜单：单元格格式化
            self.svc_tree.bind("<Button-3>", self._on_svc_right_click)
    
            # 鼠标拖动多选
            self._svc_drag_data = {"start": None, "dragging": False, "moved": False}
            self.svc_tree.bind("<ButtonPress-1>", self._on_svc_drag_start, add="+")
            self.svc_tree.bind("<B1-Motion>", self._on_svc_drag_move, add="+")
            self.svc_tree.bind("<ButtonRelease-1>", self._on_svc_drag_stop, add="+")
    
            # ---- 底栏：保存目录 + 清空 + 生成 Excel ----
            sep2 = ttk.Separator(parent, orient="horizontal")
            sep2.pack(fill="x", pady=(0, 0))
            bottom_bar = ttk.Frame(parent, padding=5)
            bottom_bar.pack(fill="x", pady=(0, 5))
    
            ttk.Label(bottom_bar, text="保存目录：", font=("微软雅黑", 11)).pack(side="left")
            saved = self.load_config()
            svc_dir = saved.get("svc_save_dir", "")
            self.svc_save_dir = tk.StringVar(value=svc_dir if svc_dir and os.path.isdir(svc_dir) else os.getcwd())
            self.entry_svc_save = ttk.Entry(bottom_bar, textvariable=self.svc_save_dir, width=28)
            self.entry_svc_save.pack(side="left", padx=5)
            ttk.Button(bottom_bar, text="浏览...", command=self._browse_svc_dir).pack(side="left", padx=3)
            ttk.Separator(bottom_bar, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
            ttk.Button(bottom_bar, text="➕ 添加行", command=self._add_empty_row).pack(side="left", padx=2)
            ttk.Button(bottom_bar, text="📌 插入行", command=self._insert_row_above).pack(side="left", padx=2)
            ttk.Button(bottom_bar, text="🗑 删除选中行", command=self._delete_selected_row).pack(side="left", padx=2)
            ttk.Button(bottom_bar, text="🔄 清空", command=self._clear_service_data).pack(side="left", padx=2)
            ttk.Separator(bottom_bar, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
            ttk.Button(bottom_bar, text="📊 生成 Excel 文档", command=self._export_service_excel,
                       style="Accent.TButton").pack(side="right", padx=5)
    
            # 加载持久化数据（无数据时预置 3 行空行）
            self._load_service_data()
            if not self.svc_tree.get_children():
                for _ in range(3):
                    self._add_empty_row(save=False)
    
        def _browse_svc_dir(self):
            folder = filedialog.askdirectory(title="选择服务记录保存目录")
            if folder:
                self.svc_save_dir.set(folder)
                self.save_config()
    
        def add_service_row(self, time_val="", item_val="", content_val="", mode_val="", remark_val=""):
            """外部调用：向服务记录表格添加一行（预通报/实况通报自动填入时使用）"""
            items = self.svc_tree.get_children()
            next_num = len(items) + 1
            row_vals = (str(next_num), time_val, item_val, content_val, mode_val, remark_val)
            self.svc_tree.insert("", tk.END, values=row_vals)
            self._save_service_data()
    
        def _renumber_rows(self):
            """重新编号所有行"""
            for i, iid in enumerate(self.svc_tree.get_children(), 1):
                vals = list(self.svc_tree.item(iid, "values"))
                vals[0] = str(i)
                self.svc_tree.item(iid, values=vals)
    
        def _add_empty_row(self, save=True):
            """在末尾添加一个空白行"""
            items = self.svc_tree.get_children()
            next_num = len(items) + 1
            self.svc_tree.insert("", tk.END, values=(str(next_num), "", "", "", "", ""))
            if save:
                self._save_service_data()
    
        def _insert_row_above(self):
            """在选中行上方插入空白行"""
            sel = self.svc_tree.selection()
            if sel:
                # 在选中行之前插入
                before_iid = sel[0]
                idx = self.svc_tree.index(before_iid)
                self.svc_tree.insert("", idx, values=("", "", "", "", "", ""))
            else:
                # 没有选中行则添加到末尾
                self.svc_tree.insert("", tk.END, values=("", "", "", "", "", ""))
            self._renumber_rows()
            self._save_service_data()
    
        def _delete_selected_row(self):
            """删除选中行"""
            sel = self.svc_tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先点击选中要删除的行")
                return
            # 先关闭正在编辑的控件
            self._svc_close_editor()
            for iid in sel:
                self.svc_tree.delete(iid)
            self._renumber_rows()
            self._save_service_data()
    
        def _svc_close_editor(self):
            """关闭当前活跃的编辑控件"""
            w = getattr(self, '_svc_active_widget', None)
            if w is not None and w.winfo_exists():
                w.destroy()
            self._svc_active_widget = None
            self._svc_editing = False
    
        def _edit_service_cell(self, event):
            """单击单元格编辑"""
            # 先关闭旧编辑器
            self._svc_close_editor()
    
            region = self.svc_tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            col_id = self.svc_tree.identify_column(event.x)
            item_id = self.svc_tree.identify_row(event.y)
            if not item_id:
                return
            col_idx = int(col_id.replace("#", "")) - 1
            col_name = self.svc_tree["columns"][col_idx]
    
            if col_name == "序号":
                return  # 序号不可编辑
    
            self._svc_editing = True  # 编辑中，防止重复触发
    
            cur_values = list(self.svc_tree.item(item_id, "values"))
            cur_text = cur_values[col_idx] if col_idx < len(cur_values) else ""
    
            # 先选中该行（视觉反馈）
            self.svc_tree.selection_set(item_id)
    
            if col_name == "事项":
                self._edit_with_combobox(item_id, col_idx, cur_text, self.service_items)
            elif col_name == "服务方式":
                self._edit_with_checkboxes(item_id, col_idx, cur_text)
            else:
                self._edit_with_entry(item_id, col_idx, cur_text, col_name)
    
        def _on_svc_double_click(self, event):
            """双击：表头编辑 或 单元格编辑"""
            # 如果正在编辑中，忽略（防止重复编辑）
            if getattr(self, '_svc_editing', False):
                return
            region = self.svc_tree.identify_region(event.x, event.y)
            if region == "heading":
                col_id = self.svc_tree.identify_column(event.x)
                if not col_id:
                    return
                ci = int(col_id.replace("#", "")) - 1
                col_name = self.svc_tree["columns"][ci]
                # 定位表头位置
                x = 0
                for c in range(ci):
                    x += self.svc_tree.column(f"#{c + 1}", "width")
                w = self.svc_tree.column(f"#{ci + 1}", "width")
                e = ttk.Entry(self.svc_tree, width=max(w // 10, 10),
                             font=("微软雅黑", 11, "bold"))
                e.place(x=x + 2, y=0, width=w - 4, height=30)
                e.insert(0, self.svc_tree.heading(col_name, "text"))
                e.lift()
                e.focus_set()
                def _save_header():
                    new_text = e.get().strip() or col_name
                    self.svc_tree.heading(col_name, text=new_text)
                    e.destroy()
                e.bind("<Return>", lambda ev: _save_header())
                e.bind("<FocusOut>", lambda ev: self.svc_tree.after(100, _save_header))
            elif region == "cell":
                # 回退到单击编辑逻辑
                self._edit_service_cell(event)
    
        def _on_svc_right_click(self, event):
            """右键菜单：单元格格式化 + 行操作"""
            item_id = self.svc_tree.identify_row(event.y)
            col_id = self.svc_tree.identify_column(event.x)
    
            menu = tk.Menu(self.root, tearoff=0)
    
            # 单元格操作（仅当点击在有效单元格上时显示）
            has_cell = False
            if col_id and item_id:
                ci = int(col_id.replace("#", "")) - 1
                if ci >= 0:
                    has_cell = True
                    menu.add_command(label="🎨 填充背景色...",
                        command=lambda iid=item_id, c=ci: self._svc_cell_bgcolor(iid, c))
                    menu.add_command(label="✏️ 文字颜色...",
                        command=lambda iid=item_id, c=ci: self._svc_cell_fgcolor(iid, c))
                    menu.add_command(label="𝐁 加粗",
                        command=lambda iid=item_id: self._svc_cell_bold(iid))
                    menu.add_separator()
                    menu.add_command(label="📝 编辑单元格",
                        command=lambda: self._edit_service_cell(event))
    
            # 行操作（始终可用）
            if has_cell:
                menu.add_separator()
            menu.add_command(label="📌 在上方插入行",
                command=self._insert_row_above)
            sel = self.svc_tree.selection()
            if len(sel) >= 2:
                menu.add_separator()
                menu.add_command(label="🔗 合并选中行",
                    command=lambda s=sel: self._svc_merge_rows(s))
            menu.post(event.x_root, event.y_root)
    
        def _svc_cell_bgcolor(self, iid, ci):
            """设置单元格背景色"""
            try:
                from tkinter import colorchooser
                c = colorchooser.askcolor(title="选择背景色")
                if c and c[1]:
                    tag_name = f"svc_bg_{c[1].replace('#', '')}_{iid}"
                    self.svc_tree.tag_configure(tag_name, background=c[1])
                    cur_tags = list(self.svc_tree.item(iid, "tags"))
                    # 移除旧背景 tag
                    cur_tags = [t for t in cur_tags if not t.startswith("svc_bg_")]
                    cur_tags.append(tag_name)
                    self.svc_tree.item(iid, tags=tuple(cur_tags))
            except Exception:
                pass
    
        def _svc_cell_fgcolor(self, iid, ci):
            """设置单元格文字颜色"""
            try:
                from tkinter import colorchooser
                c = colorchooser.askcolor(title="选择文字颜色")
                if c and c[1]:
                    tag_name = f"svc_fg_{c[1].replace('#', '')}_{iid}"
                    self.svc_tree.tag_configure(tag_name, foreground=c[1])
                    cur_tags = list(self.svc_tree.item(iid, "tags"))
                    cur_tags = [t for t in cur_tags if not t.startswith("svc_fg_")]
                    cur_tags.append(tag_name)
                    self.svc_tree.item(iid, tags=tuple(cur_tags))
            except Exception:
                pass
    
        def _svc_cell_bold(self, iid):
            """切换行加粗"""
            try:
                cur_tags = list(self.svc_tree.item(iid, "tags"))
                if "svc_bold" in cur_tags:
                    cur_tags.remove("svc_bold")
                else:
                    self.svc_tree.tag_configure("svc_bold", font=("微软雅黑", 11, "bold"))
                    cur_tags.append("svc_bold")
                self.svc_tree.item(iid, tags=tuple(cur_tags))
            except Exception:
                pass
    
        def _svc_merge_rows(self, selection):
            """垂直合并选中行：将下面行的内容追加到第一行"""
            if len(selection) < 2:
                return
            iids = list(selection)
            first_iid = iids[0]
            first_vals = list(self.svc_tree.item(first_iid, "values"))
            for iid in iids[1:]:
                vals = list(self.svc_tree.item(iid, "values"))
                for ci in range(min(len(first_vals), len(vals))):
                    if vals[ci].strip():
                        sep = "\n" if first_vals[ci].strip() else ""
                        first_vals[ci] = first_vals[ci] + sep + vals[ci]
                self.svc_tree.delete(iid)
            self.svc_tree.item(first_iid, values=first_vals)
            self._renumber_rows()
            self._save_service_data()
    
        def _on_svc_drag_start(self, event):
            """鼠标拖动开始（仅做记录，不改变选择）"""
            item_id = self.svc_tree.identify_row(event.y)
            if item_id:
                self._svc_drag_data["start"] = item_id
                self._svc_drag_data["dragging"] = True
                self._svc_drag_data["moved"] = False
    
        def _on_svc_drag_move(self, event):
            """鼠标拖动中"""
            if not self._svc_drag_data.get("dragging"):
                return
            item_id = self.svc_tree.identify_row(event.y)
            if item_id and item_id != self._svc_drag_data.get("start"):
                self._svc_drag_data["moved"] = True
                sel = set(self.svc_tree.selection())
                sel.add(item_id)
                all_items = self.svc_tree.get_children()
                start_iid = self._svc_drag_data["start"]
                try:
                    start_idx = all_items.index(start_iid)
                    cur_idx = all_items.index(item_id)
                except ValueError:
                    return
                lo, hi = min(start_idx, cur_idx), max(start_idx, cur_idx)
                for i in range(lo, hi + 1):
                    sel.add(all_items[i])
                self.svc_tree.selection_set(list(sel))
                self._svc_drag_data["start"] = item_id
    
        def _on_svc_drag_stop(self, event):
            """鼠标拖动结束"""
            self._svc_drag_data["dragging"] = False
            self._svc_drag_data["start"] = None
    
        def _on_svc_click_or_edit(self, event):
            """ButtonRelease 时：如果没有拖动，执行单元格编辑"""
            if self._svc_drag_data.get("moved"):
                self._svc_drag_data["moved"] = False
                return  # 拖动选择，不编辑
            # 防止重复编辑（编辑进行中时不再触发）
            if getattr(self, '_svc_editing', False):
                return
            self._edit_service_cell(event)
    
        def _edit_with_entry(self, item_id, col_idx, cur_text, col_name):
            """Entry / Text 编辑（内容列用 Text 支持换行）"""
            x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")
    
            if col_name == "内容":
                # 使用多行 Text 控件，支持自动换行
                text_w = tk.Text(self.svc_tree, width=max(w // 8, 20), height=5,
                                font=("微软雅黑", 11), wrap=tk.WORD,
                                padx=4, pady=4, relief="solid", borderwidth=1)
                text_w.place(x=x, y=y, width=w, height=max(h * 3, 80))
                text_w.insert("1.0", cur_text)
                text_w.lift()  # 确保在最上层
                text_w.focus_set()
                widget = text_w
                self._svc_active_widget = widget
    
                def _get_val():
                    return text_w.get("1.0", "end-1c").strip()
    
                def _delayed_destroy():
                    if text_w.winfo_exists():
                        text_w.destroy()
            else:
                entry = ttk.Entry(self.svc_tree, width=max(w // 10, 10))
                entry.place(x=x, y=y, width=w, height=h)
                entry.insert(0, cur_text)
                entry.lift()
                entry.focus_set()
                widget = entry
                self._svc_active_widget = widget
    
                def _get_val():
                    return entry.get().strip()
    
                def _delayed_destroy():
                    if entry.winfo_exists():
                        entry.destroy()
    
            destroyed = [False]
    
            def _save():
                if destroyed[0]:
                    return
                destroyed[0] = True
                self._svc_editing = False
                new_val = _get_val()
                values = list(self.svc_tree.item(item_id, "values"))
                values[col_idx] = new_val
                self.svc_tree.item(item_id, values=values)
                _delayed_destroy()
                self._save_service_data()
                # 清理全局点击绑定
                try:
                    self.root.unbind('<Button-1>', _bind_id)
                except Exception:
                    pass
    
            def _on_root_click(event):
                if destroyed[0]:
                    return
                # 判断点击是否在编辑控件内部
                w = event.widget
                while w is not None and w != self.root:
                    if w == widget:
                        return
                    try:
                        w = w.master
                    except Exception:
                        break
                self.svc_tree.after(50, _save)
    
            _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')
    
            widget.bind("<Return>", lambda e: _save())
            widget.bind("<FocusOut>", lambda e: self.svc_tree.after(150, _save))
    
        def _edit_with_combobox(self, item_id, col_idx, cur_text, options):
            """事项列 Combobox 编辑（简洁自包含的智能下拉）"""
            x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")
            if not h:  # bbox 失败则放弃
                return
    
            # ── 编辑框（只读，点击/输入触发下拉）──
            cb = ttk.Combobox(self.svc_tree, values=options, width=max(w // 10, 15))
            cb.configure(state="normal")
            cb.place(x=x, y=y, width=max(w, 180), height=h + 20)
            cb.set(cur_text)
            cb.lift()
            cb.focus_set()
            self._svc_active_widget = cb
    
            destroyed = [False]
            popup = [None]
            listbox = [None]
            timer = [None]
    
            def _cleanup():
                if popup[0] and popup[0].winfo_exists():
                    popup[0].destroy()
                popup[0] = None
                listbox[0] = None
    
            def _save_and_close():
                if destroyed[0]:
                    return
                destroyed[0] = True
                self._svc_editing = False
                _cleanup()
                new_val = cb.get().strip()
                values = list(self.svc_tree.item(item_id, "values"))
                values[col_idx] = new_val
                self.svc_tree.item(item_id, values=values)
                if cb.winfo_exists():
                    cb.destroy()
                self._save_service_data()
                try:
                    self.root.unbind('<Button-1>', _bind_id)
                except Exception:
                    pass
    
            # ── 自定义弹出列表 ──
            def _show_popup():
                _cleanup()
                vals = cb['values']
                if not vals:
                    return
                popup[0] = tk.Toplevel(self.root)
                popup[0].overrideredirect(True)
                popup[0].attributes('-topmost', True)
                popup[0].configure(bg="#3498db")
                inner = tk.Frame(popup[0], bg="white")
                inner.pack(padx=1, pady=1)
                listbox[0] = tk.Listbox(inner, font=("微软雅黑", 11),
                                        selectmode=tk.SINGLE, exportselection=False,
                                        bg="white", fg="#2c3e50",
                                        selectbackground="#3498db", selectforeground="white",
                                        relief="flat", borderwidth=0, highlightthickness=0)
                for v in vals:
                    listbox[0].insert(tk.END, v)
                n = min(len(vals), 8)
                listbox[0].config(height=n)
                listbox[0].pack(fill="both", expand=True)
                popup[0].update_idletasks()
                px = cb.winfo_rootx()
                py = cb.winfo_rooty() + cb.winfo_height()
                pw = cb.winfo_width()
                popup[0].geometry(f"{pw}x{listbox[0].winfo_reqheight() + 2}+{px}+{py}")
                popup[0].deiconify()
    
                def _on_list_click(event):
                    if listbox[0] and listbox[0].curselection():
                        text = listbox[0].get(listbox[0].curselection()[0])
                        cb.set(text)
                        cb.icursor(len(text))
                        _save_and_close()
    
                def _on_list_key(event):
                    if event.keysym == "Return":
                        _on_list_click(event)
                    elif event.keysym == "Escape":
                        _cleanup()
                        cb.focus_set()
                listbox[0].bind('<ButtonRelease-1>', _on_list_click)
                listbox[0].bind('<KeyRelease>', _on_list_key)
    
            # ── 键盘输入 → 过滤 + 弹窗 ──
            def _on_keyrelease(event):
                if destroyed[0]:
                    return
                if event.keysym in ("Escape",):
                    _cleanup()
                    return
                if event.keysym == "Return":
                    # 如果弹窗可见且有选中项，选它
                    if popup[0] and popup[0].winfo_viewable() and listbox[0]:
                        sel = listbox[0].curselection()
                        if sel:
                            cb.set(listbox[0].get(sel[0]))
                    _save_and_close()
                    return
                if event.keysym in ("Up", "Down", "Prior", "Next"):
                    # 确保弹窗显示
                    if not popup[0] or not popup[0].winfo_viewable():
                        cb['values'] = list(options)
                        _show_popup()
                    if listbox[0] and listbox[0].size():
                        listbox[0].focus_set()
                        # 移动选择
                        sel = listbox[0].curselection()
                        cur_idx = sel[0] if sel else -1
                        if event.keysym == "Down":
                            new_idx = min(cur_idx + 1, listbox[0].size() - 1)
                        elif event.keysym == "Up":
                            new_idx = max(cur_idx - 1, 0)
                        elif event.keysym == "Prior":
                            new_idx = max(cur_idx - 5, 0)
                        else:  # Next
                            new_idx = min(cur_idx + 5, listbox[0].size() - 1)
                        listbox[0].selection_clear(0, tk.END)
                        listbox[0].selection_set(new_idx)
                        listbox[0].see(new_idx)
                        # 同步填入输入框
                        cb.set(listbox[0].get(new_idx))
                        cb.icursor(len(cb.get()))
                    return
                if event.keysym in ("Left", "Right", "Tab", "Control_L", "Control_R",
                                    "Shift_L", "Shift_R", "Alt_L", "Alt_R",
                                    "Home", "End", "Caps_Lock"):
                    return
    
                # 过滤
                value = event.widget.get()
                if value:
                    filtered = [s for s in options if value.lower() in s.lower()]
                else:
                    filtered = list(options)
                cb['values'] = filtered
                # 防抖 200ms 后弹出
                if timer[0] is not None:
                    cb.after_cancel(timer[0])
                timer[0] = cb.after(200, _show_popup)
    
            cb.bind('<KeyRelease>', _on_keyrelease)
    
            # ── 点击箭头区域 → 显示完整列表 ──
            def _on_click(event):
                if event.x > event.widget.winfo_width() - 25:
                    cb['values'] = list(options)
                    _show_popup()
                    return 'break'
    
            cb.bind('<Button-1>', _on_click)
    
            # ── 失焦延迟保存 ──
            def _on_focusout(event):
                def _check():
                    if destroyed[0]:
                        return
                    if popup[0] and popup[0].winfo_viewable():
                        # 检查：焦点是否转移到了弹窗内的 listbox？
                        fw = popup[0].winfo_containing(
                            popup[0].winfo_pointerx(),
                            popup[0].winfo_pointery())
                        if fw and (fw == listbox[0] or fw == popup[0] or
                                   (hasattr(fw, 'master') and
                                    (fw.master == popup[0] or
                                     (hasattr(fw.master, 'master') and fw.master.master == popup[0])))):
                            return  # 焦点/鼠标在弹窗上，不关
                        # 额外检查：焦点控件是否是 listbox
                        focus_w = popup[0].focus_get()
                        if focus_w and (focus_w == listbox[0]):
                            return
                    self.svc_tree.after(100, _save_and_close)
                self.svc_tree.after(150, _check)
    
            cb.bind('<FocusOut>', _on_focusout)
    
            # ── 全局点击外部 → 保存 ──
            def _on_root_click(event):
                if destroyed[0]:
                    return
                w = event.widget
                while w is not None and w != self.root:
                    if w == cb:
                        return
                    if w == popup[0]:
                        return
                    try:
                        w = w.master
                    except Exception:
                        break
                self.svc_tree.after(50, _save_and_close)
    
            _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')
    
        def _edit_with_checkboxes(self, item_id, col_idx, cur_text):
            """服务方式列多选框编辑（点击外部自动关闭，无需 grab）"""
            x, y, w, h = self.svc_tree.bbox(item_id, column=f"#{col_idx + 1}")
            selected = set(s.strip() for s in cur_text.split("、") if s.strip())
    
            popup = tk.Toplevel(self.root)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)  # 始终置顶，不被其他控件遮挡
            popup.configure(bg="#3498db")
            inner = tk.Frame(popup, bg="white", padx=10, pady=10)
            inner.pack(padx=1, pady=1)
            ttk.Label(inner, text="选择服务方式", font=("微软雅黑", 12, "bold")).pack(anchor="w", pady=(0, 8))
    
            vars_ = {}
            for mode in self.service_modes:
                v = tk.BooleanVar(value=mode in selected)
                cb = tk.Checkbutton(inner, text=mode, variable=v,
                                   font=("微软雅黑", 11),
                                   bg="white", anchor="w",
                                   selectcolor="white",
                                   activebackground="white")
                cb.pack(anchor="w", fill="x")
                vars_[mode] = v
    
            def _confirm():
                self._svc_editing = False
                result = "、".join(m for m in self.service_modes if vars_[m].get())
                values = list(self.svc_tree.item(item_id, "values"))
                values[col_idx] = result
                self.svc_tree.item(item_id, values=values)
                popup.destroy()
                self._save_service_data()
                try:
                    self.root.unbind('<Button-1>', _bind_id)
                except Exception:
                    pass
    
            btn_frame = ttk.Frame(inner)
            btn_frame.pack(fill="x", pady=(10, 0))
            ttk.Button(btn_frame, text="确定", command=_confirm).pack(side="right", padx=3)
            def _cancel():
                self._svc_editing = False
                popup.destroy()
                try:
                    self.root.unbind('<Button-1>', _bind_id)
                except Exception:
                    pass
            ttk.Button(btn_frame, text="取消", command=_cancel).pack(side="right", padx=3)
    
            # 收集弹窗内所有子 widget 用于点击检测
            popup_widgets = {popup, inner, btn_frame}
            def _collect_children(w):
                popup_widgets.add(w)
                for child in w.winfo_children():
                    _collect_children(child)
            _collect_children(inner)
    
            def _on_root_click(event):
                if not popup.winfo_exists() or not popup.winfo_viewable():
                    return
                w = event.widget
                while w is not None:
                    if w in popup_widgets:
                        return
                    try:
                        w = w.master
                    except Exception:
                        break
                # 点击弹窗外 → 等同点击"确定"
                _confirm()
    
            _bind_id = self.root.bind('<Button-1>', _on_root_click, add='+')
    
            def _on_popup_destroy(event):
                try:
                    self.root.unbind('<Button-1>', _bind_id)
                except Exception:
                    pass
            popup.bind('<Destroy>', _on_popup_destroy)
    
            popup.update_idletasks()
            px = self.svc_tree.winfo_rootx() + x
            py = self.svc_tree.winfo_rooty() + y + h
            popup.geometry(f"+{px}+{py}")
            popup.focus_set()
            # 不用 grab_set()，让外部点击可以触发 root 的 Button-1 绑定
            self.root.wait_window(popup)
    
        def _save_service_data(self):
            """持久化表格数据到 JSON"""
            items = self.svc_tree.get_children()
            data = []
            for iid in items:
                vals = self.svc_tree.item(iid, "values")
                data.append(list(vals))
            try:
                with open(self.service_data_file, "w", encoding="utf-8") as f:
                    json.dump({"date": date.today().isoformat(), "weather": self.svc_weather.get(), "rows": data},
                              f, ensure_ascii=False, indent=2)
            except Exception:
                pass
    
        def _load_service_data(self):
            """从 JSON 加载表格数据"""
            if not os.path.exists(self.service_data_file):
                return
            try:
                with open(self.service_data_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.svc_tree.delete(*self.svc_tree.get_children())
                # 恢复头部
                saved_date = data.get("date", "")
                if saved_date:
                    try:
                        d = date.fromisoformat(saved_date)
                        self.svc_year.set(str(d.year))
                        self.svc_month.set(f"{d.month:02d}")
                        self.svc_day.set(f"{d.day:02d}")
                    except Exception:
                        pass
                saved_weather = data.get("weather", "")
                if saved_weather:
                    self.svc_weather.set(saved_weather)
                # 恢复行
                for row in data.get("rows", []):
                    self.svc_tree.insert("", tk.END, values=row)
            except Exception:
                pass
    
        def _clear_service_data(self):
            """清空所有数据"""
            if messagebox.askyesno("确认清空", "确定要清空表格中的所有数据吗？此操作不可恢复。"):
                self.svc_tree.delete(*self.svc_tree.get_children())
                self._save_service_data()
    
        def _get_svc_header_title(self):
            """获取表头标题"""
            y = self.svc_year.get()
            m = self.svc_month.get()
            d = self.svc_day.get()
            w = self.svc_weather.get()
            return f"{y}年{m}月{d}日{w}天气服务记录"
    
        def _export_service_excel(self):
            """导出为 Excel (.xlsx) 文件"""
            try:
                title = self._get_svc_header_title()
                # 文件名以表头文字命名：2026年06月22日强对流天气服务记录.xlsx
                filename = f"{title}.xlsx"
                save_path = os.path.join(self.svc_save_dir.get(), filename)
    
                if os.path.exists(save_path):
                    base, ext = os.path.splitext(filename)
                    count = 2
                    while True:
                        new_name = f"{base}_{count}{ext}"
                        save_path = os.path.join(self.svc_save_dir.get(), new_name)
                        if not os.path.exists(save_path):
                            break
                        count += 1
    
                # 手动构建 xlsx（zip 包 xml）
                self._write_xlsx(save_path, title)
    
                try:
                    os.startfile(os.path.normpath(os.path.dirname(save_path)))
                except Exception:
                    pass
                messagebox.showinfo("成功", f"Excel 文件已生成：\n{save_path}")
    
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{e}")
    
        def _write_xlsx(self, path, title):
            """用 XML 构建 xlsx 文件（按模板格式）"""
            columns = ["序号", "时间", "事项", "内容", "服务方式", "备注"]
            rows_data = []
            for iid in self.svc_tree.get_children():
                rows_data.append(list(self.svc_tree.item(iid, "values")))
    
            NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    
            # --- 分析时间列：尝试转为 Excel 时间序列号 ---
            def _parse_time(val):
                """尝试将 'HHMM' 或 'HH:MM' 或 'HHMM:HHMM' 格式转为 Excel 时间序列号"""
                if not val:
                    return None, val
                val = val.strip()
                for sep in [';', '；']:
                    val = val.split(sep)[0]
                val = val.strip()
                if ':' in val:
                    parts = val.split(':')
                    if len(parts) >= 2:
                        try:
                            h, m = int(parts[0]), int(parts[1])
                            return (h * 60 + m) / 1440.0, f"{h:02d}:{m:02d}"
                        except ValueError:
                            pass
                # Try 4-digit format like "1430"
                if len(val) >= 4 and val[:4].isdigit():
                    try:
                        h, m = int(val[:2]), int(val[2:4])
                        if 0 <= h <= 23 and 0 <= m <= 59:
                            return (h * 60 + m) / 1440.0, f"{h:02d}:{m:02d}"
                    except ValueError:
                        pass
                return None, val
    
            # --- 构建 sharedStrings ---
            all_strings = []
            str_index = {}
            def _add_str(s):
                s = str(s) if s is not None else ''
                if s not in str_index:
                    str_index[s] = len(all_strings)
                    all_strings.append(s)
                return str_index[s]
    
            title_idx = _add_str(title)
            for col in columns:
                _add_str(col)
    
            # 收集数据并分析时间
            time_serials = []  # per row: Excel serial or None
            for row in rows_data:
                time_val = row[1] if len(row) > 1 else ''
                serial, _ = _parse_time(time_val)
                time_serials.append(serial)
                for cell in row:
                    _add_str(str(cell))
    
            # --- Shared Strings XML ---
            sst_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst {NS} count="{len(all_strings)}" uniqueCount="{len(all_strings)}">'
            for s in all_strings:
                # 彻底转义所有 XML 敏感字符和控制字符
                escaped = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                           .replace('"', "&quot;").replace("'", "&apos;"))
                # 移除非法 XML 控制字符（除 \t \n \r）
                escaped = ''.join(c for c in escaped if ord(c) >= 0x20 or c in '\t\n\r')
                sst_xml += f'<si><t>{escaped}</t></si>'
            sst_xml += '</sst>'
    
            # --- Styles ---
            # Fonts: 0=default(宋体11), 1=title(宋体26bold), 2=header(宋体20bold), 3=time(宋体11)
            # Borders: 0=none, 1=thin all sides
            # numFmt: 0=general, 20=h:mm, 166=yyyy/m/d h:mm
            styles_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet {NS}>'
            styles_xml += '<fonts count="4">'
            styles_xml += '<font><sz val="11"/><name val="宋体"/></font>'            # 0
            styles_xml += '<font><sz val="26"/><name val="宋体"/><b/></font>'       # 1 - title
            styles_xml += '<font><sz val="20"/><name val="宋体"/><b/></font>'       # 2 - header
            styles_xml += '<font><sz val="11"/><name val="宋体"/></font>'           # 3 - time
            styles_xml += '</fonts>'
            styles_xml += '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            styles_xml += '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>'
            styles_xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            styles_xml += '<cellXfs count="6">'
            styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'  # xf0 default
            # xf1: title (center, vertical center)
            styles_xml += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
            # xf2: header (center, thin border)
            styles_xml += '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
            # xf3: data center with border (序号, 时间, 服务方式)
            styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
            # xf4: data center with border + time format
            styles_xml += '<xf numFmtId="20" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
            # xf5: content left-align with border + wrap
            styles_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
            styles_xml += '</cellXfs>'
            styles_xml += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            styles_xml += '</styleSheet>'
    
            # --- Sheet data ---
            sheet_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet {NS} {NS_R}>'
            # Page setup: landscape A4
            sheet_xml += '<sheetPr/>'
            sheet_xml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane yOffset="2" xOffset="0" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            sheet_xml += '<sheetFormatPr defaultRowHeight="15"/>'
            # Column widths (matching template)
            col_widths = [12.67, 16, 34.13, 56.63, 27.44, 28.33]
            sheet_xml += '<cols>'
            for i, w in enumerate(col_widths):
                sheet_xml += f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
            sheet_xml += '</cols>'
            sheet_xml += '<sheetData>'
    
            # Row 1: Title (merged A1:F1, s=1)
            sheet_xml += '<row r="1" ht="32.4">'
            sheet_xml += f'<c r="A1" t="s" s="1"><v>{title_idx}</v></c>'
            for ci in range(1, 6):
                sheet_xml += f'<c r="{chr(65 + ci)}1" s="1"/>'
            sheet_xml += '</row>'
    
            # Row 2: Headers (s=2)
            sheet_xml += '<row r="2" ht="25.8">'
            for ci, col in enumerate(columns):
                si = str_index[col]
                sheet_xml += f'<c r="{chr(65 + ci)}2" t="s" s="2"><v>{si}</v></c>'
            sheet_xml += '</row>'
    
            # Data rows
            for ri, row in enumerate(rows_data, 3):
                sheet_xml += f'<row r="{ri}">'
                for ci, val in enumerate(row):
                    cell_ref = f"{chr(65 + ci)}{ri}"
                    if ci == 0:  # 序号 - center
                        si = str_index[str(val)]
                        sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                    elif ci == 1:  # 时间 - try Excel time or string
                        serial = time_serials[ri - 3]
                        if serial is not None:
                            sheet_xml += f'<c r="{cell_ref}" s="4"><v>{serial}</v></c>'
                        else:
                            si = str_index[str(val)]
                            sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                    elif ci == 2 or ci == 4:  # 事项, 服务方式 - center
                        si = str_index[str(val)]
                        sheet_xml += f'<c r="{cell_ref}" t="s" s="3"><v>{si}</v></c>'
                    elif ci == 3 or ci == 5:  # 内容, 备注 - left wrap
                        si = str_index[str(val)]
                        sheet_xml += f'<c r="{cell_ref}" t="s" s="5"><v>{si}</v></c>'
                sheet_xml += '</row>'
    
            sheet_xml += '</sheetData>'
            sheet_xml += f'<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>'
            sheet_xml += f'<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.3" footer="0.3"/>'
            sheet_xml += f'<pageSetup orientation="landscape" paperSize="9" fitToHeight="1" fitToWidth="1"/>'
            sheet_xml += '</worksheet>'
    
            # --- Workbook ---
            workbook_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {NS} {NS_R}><sheets><sheet name="服务记录" sheetId="1" r:id="rId1"/></sheets></workbook>'
    
            # --- Package relationships (_rels/.rels) → 指向 xl/workbook.xml ---
            rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    
            # --- Workbook relationships (xl/_rels/workbook.xml.rels) → 指向内部部件 ---
            wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    
            # --- [Content_Types].xml ---
            content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    
            # --- Write ZIP ---
            with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.writestr('[Content_Types].xml', content_types)
                zf.writestr('_rels/.rels', rels_xml)
                zf.writestr('xl/workbook.xml', workbook_xml)
                zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
                zf.writestr('xl/sharedStrings.xml', sst_xml)
                zf.writestr('xl/styles.xml', styles_xml)
                zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
    
        # ========== 实况通报表单 ==========
