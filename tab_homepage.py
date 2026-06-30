"""首页：可编辑文本/表格 blocks 系统"""
import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from widgets import GridTable


class HomepageMixin:
        def _get_docx_path(self):
            """获取 docx 文件路径（兼容打包后）"""
            if getattr(sys, 'frozen', False):
                base = sys._MEIPASS
            else:
                base = self.app_dir
            for f in os.listdir(base):
                if f.endswith('.docx'):
                    return os.path.join(base, f)
            return None
    
        def _parse_docx(self):
            """解析 docx 返回段落列表：[{'text': str, 'font': str, 'size': int, 'bold': bool, 'align': str}]"""
            docx_path = self._get_docx_path()
            if not docx_path:
                return [{'text': '未找到文档文件，请将 .docx 文件放在程序同目录下。', 'font': '微软雅黑', 'size': 28, 'bold': False, 'align': 'left'}]
    
            NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            with zipfile.ZipFile(docx_path, 'r') as z:
                doc_xml = z.read('word/document.xml')
    
            root = ET.fromstring(doc_xml)
            body = root.find(f'.//{NS}body')
            if body is None:
                return []
    
            paragraphs = []
            for p in body.findall(f'.//{NS}p'):
                # 段落属性
                pPr = p.find(f'{NS}pPr')
                align = 'left'
                if pPr is not None:
                    jc = pPr.find(f'{NS}jc')
                    if jc is not None:
                        align = jc.get(f'{NS}val', 'left')
    
                # 收集 run
                runs = []
                for r in p.findall(f'.//{NS}r'):
                    rPr = r.find(f'{NS}rPr')
                    bold = False; font = '仿宋'; size = 28  # 默认 14pt
                    if rPr is not None:
                        b = rPr.find(f'{NS}b')
                        bold = b is not None
                        rf = rPr.find(f'{NS}rFonts')
                        if rf is not None:
                            font = rf.get(f'{NS}eastAsia', '') or rf.get(f'{NS}ascii', '') or '仿宋'
                        sz = rPr.find(f'{NS}sz')
                        if sz is not None:
                            size = int(sz.get(f'{NS}val', '28'))
                    t = r.find(f'{NS}t')
                    text = t.text if t is not None and t.text else ''
                    # 保留空白 runs（如空格）
                    if t is not None and t.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve':
                        text = t.text or ''
                    runs.append({'text': text, 'bold': bold, 'font': font, 'size': size})
    
                full_text = ''.join(r['text'] for r in runs)
                if full_text.strip() or full_text:
                    # 取第一个非空 run 的格式作为段落主格式
                    main_run = next((r for r in runs if r['text'].strip()), runs[0]) if runs else {'font': '仿宋', 'size': 28, 'bold': False}
                    paragraphs.append({
                        'text': full_text,
                        'font': main_run['font'],
                        'size': main_run['size'],
                        'bold': main_run['bold'],
                        'align': align,
                        'runs': runs if any(r['bold'] != main_run['bold'] or r['font'] != main_run['font'] for r in runs) else None,
                    })
    
            return paragraphs
    
        # ========== 首页：可编辑 block 系统 ==========
        def build_homepage(self, parent):
            """构建首页：可编辑 block 列表"""
            self.homepage_blocks_file = os.path.join(self.app_dir, "homepage_blocks.json")
            self.homepage_blocks = []
            self._homepage_load()
    
            # 滚动容器
            self.home_canvas = tk.Canvas(parent, highlightthickness=0, bg="#f0f4f8")
            home_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.home_canvas.yview)
            self.home_inner = ttk.Frame(self.home_canvas)
            self.home_inner.bind("<Configure>", lambda e: self.home_canvas.configure(
                scrollregion=self.home_canvas.bbox("all")))
            cw = self.home_canvas.create_window((0, 0), window=self.home_inner, anchor="nw")
            def _resize(event):
                self.home_canvas.itemconfig(cw, width=event.width)
            self.home_canvas.bind("<Configure>", _resize)
            self.home_canvas.configure(yscrollcommand=home_scrollbar.set)
            self.home_canvas.pack(side="left", fill="both", expand=True)
            home_scrollbar.pack(side="right", fill="y")
    
            # 滚轮由全局处理器 _global_wheel_handler 统一管理
            # 渲染所有 blocks + 底部按钮
            self._homepage_render_all()
    
        def _homepage_load(self):
            """加载首页数据（优先 JSON，否则从 docx 解析为单个文本 block）"""
            if os.path.exists(self.homepage_blocks_file):
                try:
                    with open(self.homepage_blocks_file, "r", encoding="utf-8") as f:
                        self.homepage_blocks = json.load(f)
                    return
                except Exception:
                    pass
            # 从 docx 解析为单个文本 block
            paragraphs = self._parse_docx()
            header_map = {
                "影响提示": "一、影响提示", "警戒提醒": "二、警戒提醒",
                "精细预警": "三、精细预警", "分级叫应": "四、分级叫应",
                "实况通报": "五、实况通报",
            }
            lines = []
            for p in paragraphs:
                s = p['text'].strip()
                lines.append(header_map.get(s, p['text']))
            self.homepage_blocks = [{"type": "text", "content": "\n".join(lines)}]
            self._homepage_save()
    
        def _homepage_save(self):
            try:
                with open(self.homepage_blocks_file, "w", encoding="utf-8") as f:
                    json.dump(self.homepage_blocks, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
    
        # ==================== 全量渲染 ====================
        def _homepage_render_all(self):
            """清空容器，重新渲染所有 blocks + 底部添加按钮"""
            for w in self.home_inner.winfo_children():
                w.destroy()
    
            for bi, block in enumerate(self.homepage_blocks):
                self._homepage_render_block(bi, block, edit_mode=False)
    
            # 底部添加按钮（始终在最下方）
            add_bar = ttk.Frame(self.home_inner, padding=10)
            add_bar.pack(fill="x")
            ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
            btn_frame = ttk.Frame(add_bar)
            btn_frame.pack()
            ttk.Button(btn_frame, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
            ttk.Button(btn_frame, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)
    
        # ==================== 渲染单个 block ====================
        def _homepage_render_block(self, bi, block, edit_mode=False):
            """渲染单个 block（view 或 edit 模式）"""
            if block["type"] == "text":
                self._homepage_render_text_block(bi, block, edit_mode)
            elif block["type"] == "table":
                self._homepage_render_table_block(bi, block, edit_mode)
    
        # ==================== 文本 block ====================
        def _homepage_render_text_block(self, bi, block, edit_mode=False):
            frm = ttk.LabelFrame(self.home_inner, padding=8)
            frm.pack(fill="x", padx=15, pady=(8, 4))
    
            if edit_mode:
                # ---- 格式工具栏 ----
                tb = ttk.Frame(frm)
                tb.pack(fill="x", pady=(0, 5))
    
                fonts = ["宋体", "黑体", "仿宋", "微软雅黑", "楷体", "Arial"]
                font_cb = ttk.Combobox(tb, values=fonts, width=10, state="readonly")
                font_cb.set("微软雅黑")
                font_cb.pack(side="left", padx=2)
    
                sizes = [str(s) for s in [10, 11, 12, 14, 16, 18, 20, 24, 28]]
                size_cb = ttk.Combobox(tb, values=sizes, width=4, state="readonly")
                size_cb.set("12")
                size_cb.pack(side="left", padx=2)
    
                # 辅助：获取操作范围（有选区用选区，否则全文）
                def _get_range(tw):
                    try:
                        r = tw.tag_ranges(tk.SEL)
                        if r:
                            return (r[0], r[1])
                    except Exception:
                        pass
                    return ("1.0", "end-1c")
    
                # 格式切换函数
                def make_toggle(tag, text_w):
                    def _toggle():
                        try:
                            s, e = _get_range(text_w)
                            current = text_w.tag_names(s)
                            if tag in current:
                                text_w.tag_remove(tag, s, e)
                            else:
                                text_w.tag_add(tag, s, e)
                        except Exception:
                            pass
                    return _toggle
    
                def make_apply_font(cb, text_w):
                    def _apply():
                        try:
                            s, e = _get_range(text_w)
                            fname = cb.get()
                            fsize = int(size_cb.get())
                            # 移除该范围内已有的字体 tag
                            for t in text_w.tag_names(s):
                                if t.startswith("font_"):
                                    text_w.tag_remove(t, s, e)
                            tname = f"font_{fname}_{fsize}"
                            text_w.tag_configure(tname, font=(fname, fsize))
                            text_w.tag_add(tname, s, e)
                        except Exception:
                            pass
                    return _apply
    
                def make_apply_color(text_w):
                    def _apply():
                        try:
                            from tkinter import colorchooser
                            c = colorchooser.askcolor(title="选择文字颜色")
                            if c and c[1]:
                                s, e = _get_range(text_w)
                                # 移除该范围内已有的颜色 tag（名称含 color_ 的）
                                for t in text_w.tag_names(s):
                                    if t.startswith("color_"):
                                        text_w.tag_remove(t, s, e)
                                # 新颜色 tag（去掉 # 避免潜在问题）
                                color_hex = c[1].replace("#", "")
                                tname = f"color_{color_hex}"
                                text_w.tag_configure(tname, foreground=c[1])
                                text_w.tag_add(tname, s, e)
                        except Exception:
                            pass
                    return _apply
    
                def make_align(align, text_w):
                    def _apply():
                        try:
                            s, e = _get_range(text_w)
                            tname = f"align_{align}"
                            text_w.tag_configure(tname, justify=align)
                            text_w.tag_add(tname, s, e)
                        except Exception:
                            pass
                    return _apply
    
                # 编辑区（先创建，供按钮引用）
                text_w = tk.Text(frm, wrap=tk.WORD, font=("微软雅黑", 12),
                                  relief="solid", borderwidth=1, padx=10, pady=10,
                                  height=4, undo=True)
                # 还原已保存的格式
                segments = block.get("segments", None)
                tag_configs = block.get("tag_configs", {})
                if segments:
                    for tname, cfg in tag_configs.items():
                        try:
                            tk_cfg = {}
                            if "font" in cfg:
                                font_str = cfg["font"]
                                tk_cfg["font"] = eval(font_str) if font_str.startswith("(") else font_str
                            if "foreground" in cfg:
                                tk_cfg["foreground"] = cfg["foreground"]
                            if "justify" in cfg:
                                tk_cfg["justify"] = cfg["justify"]
                            if "underline" in cfg:
                                tk_cfg["underline"] = True
                            text_w.tag_configure(tname, **tk_cfg)
                        except Exception:
                            pass
                    for seg in segments:
                        txt = seg["text"]
                        tags = seg.get("tags", [])
                        if tags:
                            text_w.insert(tk.END, txt, tuple(tags))
                        else:
                            text_w.insert(tk.END, txt)
                else:
                    text_w.insert("1.0", block["content"])
    
                # 配置基础 tag
                text_w.tag_configure("bold", font=("微软雅黑", 12, "bold"))
                text_w.tag_configure("italic", font=("微软雅黑", 12, "italic"))
                text_w.tag_configure("underline", underline=True)
    
                # 格式按钮（放在编辑区上方）
                ttk.Button(tb, text="B", width=2, command=make_toggle("bold", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="I", width=2, command=make_toggle("italic", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="U", width=2, command=make_toggle("underline", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="A", width=2, command=make_apply_color(text_w)).pack(side="left", padx=3)
                ttk.Button(tb, text="≡", width=2, command=make_align("left", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="⊜", width=2, command=make_align("center", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="⊐", width=2, command=make_align("right", text_w)).pack(side="left", padx=1)
                ttk.Button(tb, text="应用字体", width=7, command=make_apply_font(font_cb, text_w)).pack(side="left", padx=5)
    
                # 自适应高度：根据内容计算所需行数
                def _auto_height(tw=None):
                    tw = tw or text_w
                    try:
                        tw.update_idletasks()
                        result = tw.tk.call(tw._w, 'count', '-displaylines', '1.0', 'end-1c')
                        actual_lines = int(result) if result else int(tw.index("end-1c").split(".")[0])
                    except Exception:
                        actual_lines = int(tw.index("end-1c").split(".")[0])
                    tw.configure(height=max(actual_lines + 2, 6))
                text_w.pack(side="left", fill="both", expand=True)
                # 初始化自适应高度
                _auto_height(text_w)
                # 输入时动态调整高度（防抖 100ms）
                _height_timer = [None]
                def _on_text_change(event=None):
                    if _height_timer[0] is not None:
                        text_w.after_cancel(_height_timer[0])
                    _height_timer[0] = text_w.after(100, lambda: _auto_height(text_w))
                text_w.bind("<KeyRelease>", _on_text_change)
                # 窗口大小变化时也调整高度
                text_w.bind("<Configure>", lambda e: _auto_height(text_w) if e.widget == text_w and e.width > 10 else None)
    
                # 确认 / 取消按钮
                btn_row = ttk.Frame(frm)
                btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
                ttk.Button(btn_row, text="✅ 确认",
                           command=lambda w=text_w, b=bi: self._homepage_text_confirm(w, b)).pack(side="right", padx=5)
                ttk.Button(btn_row, text="❌ 取消",
                           command=lambda b=bi, orig=dict(block): self._homepage_text_cancel(b, orig)).pack(side="right", padx=5)
            else:
                # ---- 只读显示（还原格式） ----
                text_w = tk.Text(frm, wrap=tk.WORD, font=("微软雅黑", 12),
                                  bg="#f0f4f8", fg="#2c3e50",
                                  relief="flat", borderwidth=0,
                                  height=1,  # 初始最小高度，后续自动调整
                                  state="disabled")
                text_w.pack(fill="both", expand=True)
                text_w.configure(state="normal")
    
                segments = block.get("segments", None)
                tag_configs = block.get("tag_configs", {})
                if segments:
                    # 预注册所有 tag 配置
                    for tname, cfg in tag_configs.items():
                        try:
                            tk_cfg = {}
                            if "font" in cfg:
                                font_str = cfg["font"]
                                tk_cfg["font"] = eval(font_str) if font_str.startswith("(") else font_str
                            if "foreground" in cfg:
                                tk_cfg["foreground"] = cfg["foreground"]
                            if "justify" in cfg:
                                tk_cfg["justify"] = cfg["justify"]
                            if "underline" in cfg:
                                tk_cfg["underline"] = True
                            text_w.tag_configure(tname, **tk_cfg)
                        except Exception:
                            pass
                    for seg in segments:
                        txt = seg["text"]
                        tags = seg.get("tags", [])
                        if tags:
                            text_w.insert(tk.END, txt, tuple(tags))
                        else:
                            text_w.insert(tk.END, txt)
                else:
                    text_w.insert("1.0", block["content"])
                text_w.configure(state="disabled")
                # 自动调整高度：用 Tcl count -displaylines 计算含自动换行的真实行数
                _view_height_busy = [False]
                def _view_auto_height(tw=text_w):
                    if _view_height_busy[0]:
                        return
                    _view_height_busy[0] = True
                    try:
                        tw.update_idletasks()
                        try:
                            result = tw.tk.call(tw._w, 'count', '-displaylines', '1.0', 'end-1c')
                            actual_lines = int(result) if result else int(tw.index("end-1c").split(".")[0])
                        except Exception:
                            actual_lines = int(tw.index("end-1c").split(".")[0])
                        new_h = max(actual_lines + 1, 4)
                        if int(tw.cget("height")) != new_h:
                            tw.configure(height=new_h)
                    finally:
                        _view_height_busy[0] = False
                _view_auto_height()
                text_w.bind("<Configure>", lambda e, tw=text_w: (None if e.widget != tw or e.width <= 10
                    else tw.after(80, _view_auto_height)))
    
                # ---- 拖拽调整高度把手 ----
                grip = tk.Frame(frm, height=6, cursor="sb_v_double_arrow", bg="#c0c8d0")
                grip.pack(fill="x", pady=(2, 0))
                _drag_grip = [False, 0, 0]  # [dragging, start_y, start_height]
                def _grip_press(event):
                    _drag_grip[0] = True
                    _drag_grip[1] = event.y_root
                    _drag_grip[2] = int(text_w.cget("height"))
                def _grip_move(event):
                    if not _drag_grip[0]:
                        return
                    dy = event.y_root - _drag_grip[1]
                    new_h = max(3, _drag_grip[2] + dy // 16)
                    if int(text_w.cget("height")) != new_h:
                        text_w.configure(height=new_h)
                        # 取消自动高度（手动设置优先）
                        text_w.unbind("<Configure>")
                def _grip_stop(event):
                    _drag_grip[0] = False
                    # 保存手动高度到 block 并持久化
                    block["_manual_height"] = int(text_w.cget("height"))
                    self._homepage_save()
                grip.bind("<ButtonPress-1>", _grip_press)
                grip.bind("<B1-Motion>", _grip_move)
                grip.bind("<ButtonRelease-1>", _grip_stop)
                # 恢复手动高度
                if block.get("_manual_height"):
                    try:
                        text_w.configure(height=max(int(block["_manual_height"]), 3))
                    except Exception:
                        pass
    
                # 编辑 / 删除按钮
                btn_row = ttk.Frame(frm)
                btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
                ttk.Button(btn_row, text="✏️ 编辑",
                           command=lambda b=bi: self._homepage_edit_block(b)).pack(side="right", padx=5)
                ttk.Button(btn_row, text="🗑 删除",
                           command=lambda b=bi: self._homepage_delete_block(b)).pack(side="right", padx=5)
    
        def _homepage_text_confirm(self, text_w, bi):
            """确认文本编辑：保存文本 + 格式信息"""
            content = text_w.get("1.0", "end-1c")
            if not content.strip():
                del self.homepage_blocks[bi]
                self._homepage_save()
                self._homepage_render_all()
                return
    
            # 使用 dump 提取完整格式信息
            # dump 返回: [('text', 'H', []), ('tagon', 'bold', {}), ('text', 'e', ['bold']), ...]
            dumped = text_w.dump("1.0", "end-1c", tag=True, text=True)
            # 简化存储：逐段记录 [text, active_tags, tag_configs]
            segments = []
            current_text = ""
            current_tags = set()
            all_tag_configs = {}
    
            for item in dumped:
                if item[0] == "text":
                    current_text += item[1]
                elif item[0] == "tagon":
                    tname = item[1]
                    if current_text:
                        segments.append({"text": current_text, "tags": list(current_tags)})
                        current_text = ""
                    current_tags.add(tname)
                    # 记录 tag 配置
                    if tname not in all_tag_configs:
                        cfg = {}
                        try:
                            fi = text_w.tag_cget(tname, "font")
                            if fi:
                                cfg["font"] = str(fi)
                        except Exception:
                            pass
                        try:
                            fg = text_w.tag_cget(tname, "foreground")
                            if fg:
                                cfg["foreground"] = str(fg)
                        except Exception:
                            pass
                        try:
                            jf = text_w.tag_cget(tname, "justify")
                            if jf and jf != "left":
                                cfg["justify"] = str(jf)
                        except Exception:
                            pass
                        try:
                            ul = text_w.tag_cget(tname, "underline")
                            if ul and ul != "0":
                                cfg["underline"] = "1"
                        except Exception:
                            pass
                        if cfg:
                            all_tag_configs[tname] = cfg
                elif item[0] == "tagoff":
                    tname = item[1]
                    if current_text:
                        segments.append({"text": current_text, "tags": list(current_tags)})
                        current_text = ""
                    current_tags.discard(tname)
    
            if current_text:
                segments.append({"text": current_text, "tags": list(current_tags) if current_tags else []})
    
            self.homepage_blocks[bi]["content"] = content
            self.homepage_blocks[bi]["segments"] = segments
            self.homepage_blocks[bi]["tag_configs"] = all_tag_configs
            self._homepage_save()
            self._homepage_render_all()
    
        def _homepage_text_cancel(self, bi, original):
            """取消文本编辑：恢复原始完整状态"""
            self.homepage_blocks[bi] = original
            self._homepage_save()
            self._homepage_render_all()
    
        # ==================== 表格 block ====================
        # ==================== 网格表格（Entry 网格） ====================
        # ---- _homepage_render_table_block（重写为网格表格） ----
        def _homepage_render_table_block(self, bi, block, edit_mode=False):
            frm = ttk.LabelFrame(self.home_inner, padding=8)
            frm.pack(fill="x", padx=15, pady=(8, 4))
    
            headers = list(block.get("headers", ["列1", "列2", "列3"]))
            rows = block.get("rows", [["", "", ""]])
            for r in rows:
                while len(r) < len(headers):
                    r.append("")
            saved_styles = block.get("cell_styles", {})
            cell_styles = {}
            for key, style in saved_styles.items():
                parts = key.split(",")
                if len(parts) == 2:
                    try: cell_styles[(int(parts[0]), int(parts[1]))] = style
                    except Exception: pass
    
            gt = GridTable(frm, headers, rows, cell_styles, edit_mode)
            gt.frame.pack(fill="both", expand=True)
    
            if edit_mode:
                op_row = ttk.Frame(frm)
                op_row.pack(fill="x", pady=(5, 0))
                ttk.Label(op_row, text="行：", font=("微软雅黑", 10)).pack(side="left", padx=(0, 2))
                ttk.Button(op_row, text="➕ 末尾", command=gt.add_row_at_end).pack(side="left", padx=1)
                ttk.Button(op_row, text="📌 插入", command=gt.insert_row_above).pack(side="left", padx=1)
                ttk.Button(op_row, text="🗑 删除", command=gt.delete_selected_rows).pack(side="left", padx=1)
                ttk.Separator(op_row, orient="vertical").pack(side="left", fill="y", padx=5, pady=2)
                ttk.Label(op_row, text="列：", font=("微软雅黑", 10)).pack(side="left", padx=(0, 2))
                ttk.Button(op_row, text="➕ 末尾", command=gt.add_column_at_end).pack(side="left", padx=1)
                ttk.Button(op_row, text="🗑 末尾", command=gt.delete_last_column).pack(side="left", padx=1)
    
                btn_row = ttk.Frame(frm)
                btn_row.pack(side="bottom", anchor="e", pady=(5, 0))
                def _confirm():
                    hdrs, data_rows, styles = gt.get_data()
                    block["headers"] = hdrs
                    block["rows"] = data_rows
                    block["cell_styles"] = styles
                    self._homepage_save()
                    self._homepage_render_all()
                ttk.Button(btn_row, text="✅ 确认", command=_confirm).pack(side="right", padx=5)
                ttk.Button(btn_row, text="❌ 取消",
                           command=lambda: self._homepage_render_all()).pack(side="right", padx=5)
            else:
                btn_row = ttk.Frame(frm)
                btn_row.pack(fill="x", pady=(5, 0))
                ttk.Button(btn_row, text="✏️ 编辑此表格",
                           command=lambda b=bi: self._homepage_edit_block(b)).pack(side="left", padx=2)
                ttk.Button(btn_row, text="🗑 删除此表格",
                           command=lambda b=bi: self._homepage_delete_block(b)).pack(side="right", padx=2)
    
        def _homepage_edit_block(self, bi):
            """进入编辑模式：只渲染该 block 为编辑态"""
            # 清除容器
            for w in self.home_inner.winfo_children():
                w.destroy()
            for i, blk in enumerate(self.homepage_blocks):
                self._homepage_render_block(i, blk, edit_mode=(i == bi))
            # 底部按钮
            add_bar = ttk.Frame(self.home_inner, padding=10)
            add_bar.pack(fill="x")
            ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
            ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
            ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)
    
        def _homepage_delete_block(self, bi):
            """删除 block（带确认弹窗）"""
            popup = tk.Toplevel(self.root)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            popup.configure(bg="#e74c3c")
            inner = tk.Frame(popup, bg="white", padx=20, pady=18)
            inner.pack(padx=2, pady=2)
            ttk.Label(inner, text="确定要删除吗？", font=("微软雅黑", 12, "bold")).pack(pady=(0, 12))
            bf = ttk.Frame(inner)
            bf.pack()
            def _do():
                popup.destroy()
                del self.homepage_blocks[bi]
                self._homepage_save()
                self._homepage_render_all()
                try:
                    self.root.unbind('<Button-1>', _bid)
                except Exception:
                    pass
            def _cancel():
                popup.destroy()
                try:
                    self.root.unbind('<Button-1>', _bid)
                except Exception:
                    pass
            ttk.Button(bf, text="确定", command=_do).pack(side="left", padx=8)
            ttk.Button(bf, text="取消", command=_cancel).pack(side="left", padx=8)
            # 收集弹窗控件
            pw = {popup, inner, bf}
            def _coll(w):
                pw.add(w)
                for c in w.winfo_children():
                    _coll(c)
            _coll(inner)
            def _click(event):
                if not popup.winfo_exists():
                    return
                w = event.widget
                while w:
                    if w in pw:
                        return
                    try:
                        w = w.master
                    except Exception:
                        break
                _cancel()
            _bid = self.root.bind('<Button-1>', _click, add='+')
            popup.update_idletasks()
            px = self.root.winfo_rootx() + self.root.winfo_width() // 2 - 80
            py = self.root.winfo_rooty() + self.root.winfo_height() // 2 - 50
            popup.geometry(f"+{px}+{py}")
            popup.focus_set()
    
        def _homepage_add_text(self):
            """添加文本 block 并进入编辑"""
            self.homepage_blocks.append({"type": "text", "content": ""})
            self._homepage_save()
            bi = len(self.homepage_blocks) - 1
            # 渲染：其他 view，新 block edit
            for w in self.home_inner.winfo_children():
                w.destroy()
            for i, blk in enumerate(self.homepage_blocks):
                self._homepage_render_block(i, blk, edit_mode=(i == bi))
            add_bar = ttk.Frame(self.home_inner, padding=10)
            add_bar.pack(fill="x")
            ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
            ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
            ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)
    
        def _homepage_add_table(self):
            """添加表格 block 并进入编辑"""
            self.homepage_blocks.append({
                "type": "table",
                "headers": ["列1", "列2", "列3"],
                "rows": [["", "", ""], ["", "", ""], ["", "", ""]],
            })
            self._homepage_save()
            bi = len(self.homepage_blocks) - 1
            for w in self.home_inner.winfo_children():
                w.destroy()
            for i, blk in enumerate(self.homepage_blocks):
                self._homepage_render_block(i, blk, edit_mode=(i == bi))
            add_bar = ttk.Frame(self.home_inner, padding=10)
            add_bar.pack(fill="x")
            ttk.Separator(add_bar, orient="horizontal").pack(fill="x", pady=(5, 10))
            ttk.Button(add_bar, text="➕ 添加文本", command=self._homepage_add_text).pack(side="left", padx=10)
            ttk.Button(add_bar, text="➕ 添加表格", command=self._homepage_add_table).pack(side="left", padx=10)
    
        # ========== 叫应名单（解析 docx 含表格） ==========
