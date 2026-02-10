# -*- coding: utf-8 -*-
"""
库存自动化控制面板 - 傻瓜式一体化界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from pathlib import Path
import sys
import shutil
import yaml
import threading
import pyautogui


def get_base_dir():
    """获取程序基础目录，兼容打包后的exe"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的exe
        return Path(sys.executable).parent
    else:
        # 开发环境
        return Path(__file__).parent


class ControlPanel:
    """主控制面板"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("库存批量修改自动化工具")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # 路径配置
        self.base_dir = get_base_dir()
        self.assets_dir = self.base_dir / "assets"
        self.data_dir = self.base_dir / "data"
        self.config_path = self.base_dir / "config.yaml"

        # 确保目录存在
        self.assets_dir.mkdir(exist_ok=True)
        self.data_dir.mkdir(exist_ok=True)

        # 加载配置
        self.config = self.load_config()
        self.steps = self.config.get('steps', [])

        # 图片缓存
        self.image_cache = {}

        self.setup_ui()
        self.refresh_all()

    def load_config(self):
        """加载配置"""
        if self.config_path.exists():
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        return self.get_default_config()

    def get_default_config(self):
        """默认配置"""
        return {
            'excel': {
                'file_path': '',
                'code_column': '编码',
                'quantity_column': '库存数',
                'sheet_name': None
            },
            'settings': {
                'confidence': 0.8,
                'default_wait': 0.5,
                'timeout': 10,
                'failsafe': True
            },
            'steps': []
        }

    def save_config(self):
        """保存配置"""
        self.config['steps'] = self.steps
        with open(self.config_path, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

    def setup_ui(self):
        """构建主界面"""
        # 创建notebook标签页
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 标签页1: Excel配置
        self.tab_excel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text="  1. 选择Excel  ")
        self.setup_excel_tab()

        # 标签页2: 步骤配置
        self.tab_steps = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_steps, text="  2. 配置步骤  ")
        self.setup_steps_tab()

        # 标签页3: 运行
        self.tab_run = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_run, text="  3. 开始运行  ")
        self.setup_run_tab()

    # ==================== Excel 标签页 ====================
    def setup_excel_tab(self):
        """Excel配置标签页"""
        frame = ttk.LabelFrame(self.tab_excel, text="Excel 文件设置", padding=20)
        frame.pack(fill='both', expand=True, padx=20, pady=20)

        # 文件选择
        ttk.Label(frame, text="Excel文件:", font=('', 10)).grid(row=0, column=0, sticky='e', pady=10)
        self.excel_path_var = tk.StringVar(value=self.config.get('excel', {}).get('file_path', ''))
        ttk.Entry(frame, textvariable=self.excel_path_var, width=50).grid(row=0, column=1, padx=10)
        ttk.Button(frame, text="浏览...", command=self.browse_excel).grid(row=0, column=2)

        # 列名配置
        ttk.Label(frame, text="编码列名:", font=('', 10)).grid(row=1, column=0, sticky='e', pady=10)
        self.code_col_var = tk.StringVar(value=self.config.get('excel', {}).get('code_column', '编码'))
        ttk.Entry(frame, textvariable=self.code_col_var, width=20).grid(row=1, column=1, sticky='w', padx=10)

        ttk.Label(frame, text="库存列名:", font=('', 10)).grid(row=2, column=0, sticky='e', pady=10)
        self.qty_col_var = tk.StringVar(value=self.config.get('excel', {}).get('quantity_column', '库存数'))
        ttk.Entry(frame, textvariable=self.qty_col_var, width=20).grid(row=2, column=1, sticky='w', padx=10)

        # 预览区域
        preview_frame = ttk.LabelFrame(frame, text="数据预览", padding=10)
        preview_frame.grid(row=3, column=0, columnspan=3, sticky='nsew', pady=20)
        frame.rowconfigure(3, weight=1)

        self.excel_preview = ttk.Treeview(preview_frame, height=8)
        self.excel_preview.pack(fill='both', expand=True)

        # 保存按钮
        ttk.Button(frame, text="保存Excel设置", command=self.save_excel_config).grid(row=4, column=1, pady=10)

    def browse_excel(self):
        """浏览Excel文件"""
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            # 复制到data目录
            dest = self.data_dir / Path(path).name
            if Path(path) != dest:
                shutil.copy(path, dest)
            self.excel_path_var.set(f"data/{dest.name}")
            self.preview_excel(dest)

    def preview_excel(self, path):
        """预览Excel内容"""
        try:
            import pandas as pd
            df = pd.read_excel(path, nrows=10)

            # 清空旧数据
            self.excel_preview.delete(*self.excel_preview.get_children())

            # 设置列
            self.excel_preview['columns'] = list(df.columns)
            self.excel_preview['show'] = 'headings'
            for col in df.columns:
                self.excel_preview.heading(col, text=col)
                self.excel_preview.column(col, width=100)

            # 插入数据
            for _, row in df.iterrows():
                self.excel_preview.insert('', 'end', values=list(row))
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败: {e}")

    def save_excel_config(self):
        """保存Excel配置"""
        self.config['excel'] = {
            'file_path': self.excel_path_var.get(),
            'code_column': self.code_col_var.get(),
            'quantity_column': self.qty_col_var.get(),
            'sheet_name': None
        }
        self.save_config()
        messagebox.showinfo("成功", "Excel设置已保存!")

    # ==================== 步骤配置标签页 ====================
    def setup_steps_tab(self):
        """步骤配置标签页"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.tab_steps)
        toolbar.pack(fill='x', padx=10, pady=10)

        ttk.Button(toolbar, text="添加步骤", command=self.add_step).pack(side='left', padx=5)
        ttk.Button(toolbar, text="编辑步骤", command=self.edit_step).pack(side='left', padx=5)
        ttk.Button(toolbar, text="删除步骤", command=self.delete_step).pack(side='left', padx=5)
        ttk.Button(toolbar, text="上移", command=self.move_step_up).pack(side='left', padx=5)
        ttk.Button(toolbar, text="下移", command=self.move_step_down).pack(side='left', padx=5)
        ttk.Button(toolbar, text="保存步骤", command=self.save_steps).pack(side='right', padx=5)

        # 步骤列表
        list_frame = ttk.Frame(self.tab_steps)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)

        columns = ('序号', '步骤名称', '动作类型', '目标/内容')
        self.steps_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)

        self.steps_tree.heading('序号', text='序号')
        self.steps_tree.heading('步骤名称', text='步骤名称')
        self.steps_tree.heading('动作类型', text='动作类型')
        self.steps_tree.heading('目标/内容', text='目标/内容')

        self.steps_tree.column('序号', width=50)
        self.steps_tree.column('步骤名称', width=150)
        self.steps_tree.column('动作类型', width=100)
        self.steps_tree.column('目标/内容', width=200)

        self.steps_tree.pack(fill='both', expand=True)
        self.steps_tree.bind('<Double-1>', lambda e: self.edit_step())

    def refresh_steps(self):
        """刷新步骤列表"""
        for item in self.steps_tree.get_children():
            self.steps_tree.delete(item)

        for i, step in enumerate(self.steps, 1):
            # 显示坐标或文本/按键
            if 'x' in step and 'y' in step:
                target = f"坐标({step['x']}, {step['y']})"
            else:
                target = step.get('text', step.get('key', ''))
            self.steps_tree.insert('', 'end', values=(i, step.get('name', ''), step.get('action', ''), target))

    def get_selected_step_index(self):
        """获取选中步骤索引"""
        selection = self.steps_tree.selection()
        if not selection:
            return None
        item = self.steps_tree.item(selection[0])
        return int(item['values'][0]) - 1

    def delete_step(self):
        """删除步骤"""
        idx = self.get_selected_step_index()
        if idx is not None:
            del self.steps[idx]
            self.refresh_steps()

    def move_step_up(self):
        """上移步骤"""
        idx = self.get_selected_step_index()
        if idx and idx > 0:
            self.steps[idx], self.steps[idx-1] = self.steps[idx-1], self.steps[idx]
            self.refresh_steps()

    def move_step_down(self):
        """下移步骤"""
        idx = self.get_selected_step_index()
        if idx is not None and idx < len(self.steps) - 1:
            self.steps[idx], self.steps[idx+1] = self.steps[idx+1], self.steps[idx]
            self.refresh_steps()

    def save_steps(self):
        """保存步骤配置"""
        self.save_config()
        messagebox.showinfo("成功", "步骤配置已保存!")

    def add_step(self):
        """添加步骤"""
        self.open_step_dialog()

    def edit_step(self):
        """编辑步骤"""
        idx = self.get_selected_step_index()
        if idx is not None:
            self.open_step_dialog(idx)

    def open_step_dialog(self, edit_idx=None):
        """打开步骤编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑步骤" if edit_idx is not None else "添加步骤")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()

        step_data = self.steps[edit_idx] if edit_idx is not None else {}

        # 步骤名称
        ttk.Label(dialog, text="步骤名称:", font=('', 10)).grid(row=0, column=0, padx=10, pady=8, sticky='e')
        name_var = tk.StringVar(value=step_data.get('name', ''))
        ttk.Entry(dialog, textvariable=name_var, width=35).grid(row=0, column=1, padx=10, pady=8, sticky='w')

        # 动作类型
        ttk.Label(dialog, text="动作类型:", font=('', 10)).grid(row=1, column=0, padx=10, pady=8, sticky='e')
        action_map = {
            '点击坐标': 'click',
            '双击坐标': 'double_click',
            '输入文本': 'type_text',
            '按键': 'press_key',
            '等待秒数': 'wait',
            '清空输入框': 'clear_input'
        }
        action_display = {v: k for k, v in action_map.items()}
        action_list = list(action_map.keys())
        action_combo = ttk.Combobox(dialog, values=action_list, width=32, state='readonly')
        # 设置默认选中项
        current_action_code = step_data.get('action', 'click')
        current_action_text = action_display.get(current_action_code, '点击坐标')
        if current_action_text in action_list:
            action_combo.current(action_list.index(current_action_text))
        else:
            action_combo.current(0)
        action_combo.grid(row=1, column=1, padx=10, pady=8, sticky='w')

        # 坐标输入
        ttk.Label(dialog, text="点击坐标:", font=('', 10)).grid(row=2, column=0, padx=10, pady=8, sticky='e')
        coord_frame = ttk.Frame(dialog)
        coord_frame.grid(row=2, column=1, padx=10, pady=8, sticky='w')

        ttk.Label(coord_frame, text="X:").pack(side='left')
        x_var = tk.StringVar(value=str(step_data.get('x', '')))
        ttk.Entry(coord_frame, textvariable=x_var, width=6).pack(side='left', padx=2)

        ttk.Label(coord_frame, text="Y:").pack(side='left', padx=(10,0))
        y_var = tk.StringVar(value=str(step_data.get('y', '')))
        ttk.Entry(coord_frame, textvariable=y_var, width=6).pack(side='left', padx=2)

        def get_mouse_pos():
            dialog.withdraw()
            self.root.withdraw()
            messagebox.showinfo("提示", "点击确定后，3秒内把鼠标移到目标位置")
            import time
            time.sleep(3)
            x, y = pyautogui.position()
            x_var.set(str(x))
            y_var.set(str(y))

            # 使用after延迟恢复窗口，确保焦点正确
            def restore_windows():
                self.root.deiconify()
                self.root.update()
                dialog.deiconify()
                dialog.update()
                # 再次延迟获取焦点
                def focus_dialog():
                    dialog.lift()
                    dialog.focus_force()
                    dialog.grab_set()
                    messagebox.showinfo("完成", f"已获取坐标: ({x}, {y})\n请点击【保存】按钮保存此步骤")
                dialog.after(100, focus_dialog)
            self.root.after(100, restore_windows)

        ttk.Button(coord_frame, text="获取鼠标位置", command=get_mouse_pos).pack(side='left', padx=10)

        # 输入文本
        ttk.Label(dialog, text="输入内容:", font=('', 10)).grid(row=3, column=0, padx=10, pady=8, sticky='e')
        text_var = tk.StringVar(value=step_data.get('text', ''))
        ttk.Entry(dialog, textvariable=text_var, width=35).grid(row=3, column=1, padx=10, pady=8, sticky='w')
        ttk.Label(dialog, text="提示: 用 {code} 代表编码, {quantity} 代表库存数", foreground='gray').grid(row=4, column=1, sticky='w', padx=10)

        # 按键
        ttk.Label(dialog, text="按键:", font=('', 10)).grid(row=5, column=0, padx=10, pady=8, sticky='e')
        key_var = tk.StringVar(value=step_data.get('key', ''))
        common_keys = [
            '',
            'enter',
            'tab',
            'escape',
            'backspace',
            'delete',
            'space',
            'up',
            'down',
            'left',
            'right',
            'ctrl+a',
            'ctrl+c',
            'ctrl+v',
            'ctrl+x',
            'ctrl+s',
            'ctrl+z',
            'alt+tab',
            'alt+f4',
            'f1',
            'f2',
            'f5'
        ]
        key_combo = ttk.Combobox(dialog, textvariable=key_var, values=common_keys, width=32)
        key_combo.set(step_data.get('key', ''))
        key_combo.grid(row=5, column=1, padx=10, pady=8, sticky='w')
        ttk.Label(dialog, text="可选择常用按键或手动输入其他按键", foreground='gray').grid(row=6, column=1, sticky='w', padx=10)

        # 等待时间
        ttk.Label(dialog, text="操作后等待(秒):", font=('', 10)).grid(row=7, column=0, padx=10, pady=8, sticky='e')
        wait_var = tk.StringVar(value=str(step_data.get('wait_after', 0.5)))
        ttk.Entry(dialog, textvariable=wait_var, width=10).grid(row=7, column=1, padx=10, pady=8, sticky='w')

        # 清空选项
        clear_var = tk.BooleanVar(value=step_data.get('clear_first', False))
        ttk.Checkbutton(dialog, text="输入前先清空原内容", variable=clear_var).grid(row=8, column=1, padx=10, pady=8, sticky='w')

        def save_step():
            action_text = action_combo.get()
            step = {
                'name': name_var.get(),
                'action': action_map.get(action_text, 'click')
            }

            # 坐标
            if x_var.get() and y_var.get():
                try:
                    step['x'] = int(x_var.get())
                    step['y'] = int(y_var.get())
                except ValueError:
                    pass

            # 输入文本
            if text_var.get():
                step['text'] = text_var.get()

            # 按键
            if key_var.get():
                step['key'] = key_var.get()

            # 清空选项
            if clear_var.get():
                step['clear_first'] = True

            # 等待时间
            try:
                wait = float(wait_var.get())
                if wait > 0:
                    step['wait_after'] = wait
            except ValueError:
                pass

            if edit_idx is not None:
                self.steps[edit_idx] = step
            else:
                self.steps.append(step)

            self.refresh_steps()
            dialog.destroy()

        ttk.Button(dialog, text="保存", command=save_step).grid(row=9, column=1, pady=20)

    # ==================== 运行标签页 ====================
    def setup_run_tab(self):
        """运行标签页"""
        # 状态检查
        check_frame = ttk.LabelFrame(self.tab_run, text="运行前检查", padding=15)
        check_frame.pack(fill='x', padx=20, pady=10)

        self.check_labels = {}
        checks = [
            ('excel', 'Excel文件'),
            ('steps', '操作步骤')
        ]
        for i, (key, text) in enumerate(checks):
            ttk.Label(check_frame, text=f"{text}:").grid(row=i, column=0, sticky='e', padx=10, pady=5)
            self.check_labels[key] = ttk.Label(check_frame, text="检查中...")
            self.check_labels[key].grid(row=i, column=1, sticky='w', padx=10, pady=5)

        # 执行次数设置
        limit_frame = ttk.Frame(self.tab_run)
        limit_frame.pack(pady=10)

        ttk.Label(limit_frame, text="循环次数:", font=('', 10)).pack(side='left')
        self.limit_var = tk.StringVar(value="0")
        ttk.Entry(limit_frame, textvariable=self.limit_var, width=8).pack(side='left', padx=5)
        ttk.Label(limit_frame, text="(0 = 按Excel数据条数)", foreground='gray').pack(side='left')

        # 运行按钮
        btn_frame = ttk.Frame(self.tab_run)
        btn_frame.pack(pady=20)

        self.run_btn = ttk.Button(btn_frame, text="开始运行", command=self.start_bot)
        self.run_btn.pack(side='left', padx=10)

        ttk.Button(btn_frame, text="刷新检查", command=self.check_ready).pack(side='left', padx=10)

        # 日志区域
        log_frame = ttk.LabelFrame(self.tab_run, text="运行日志", padding=10)
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)

        self.log_text = ScrolledText(log_frame, height=15, font=('Consolas', 9))
        self.log_text.pack(fill='both', expand=True)

    def check_ready(self):
        """检查运行条件"""
        all_ok = True

        # 检查Excel - 优先检查界面上的值
        excel_path = self.excel_path_var.get() or self.config.get('excel', {}).get('file_path', '')
        if excel_path:
            # 处理路径
            full_path = Path(excel_path)
            if not full_path.is_absolute():
                full_path = self.base_dir / excel_path
            if full_path.exists():
                self.check_labels['excel'].config(text="OK", foreground='green')
            else:
                self.check_labels['excel'].config(text=f"文件不存在: {excel_path}", foreground='red')
                all_ok = False
        else:
            self.check_labels['excel'].config(text="未配置", foreground='red')
            all_ok = False

        # 检查步骤
        if self.steps:
            self.check_labels['steps'].config(text=f"已配置 {len(self.steps)} 个步骤", foreground='green')
        else:
            self.check_labels['steps'].config(text="未配置", foreground='red')
            all_ok = False

        return all_ok

    def log(self, msg):
        """写入日志"""
        self.log_text.insert(tk.END, msg + '\n')
        self.log_text.see(tk.END)
        self.root.update()

    def start_bot(self):
        """启动自动化"""
        if not self.check_ready():
            messagebox.showwarning("提示", "请先完成所有配置!")
            return

        # 获取执行次数
        try:
            limit = int(self.limit_var.get())
        except ValueError:
            limit = 0

        # 保存配置
        self.save_config()

        self.log("=" * 40)
        self.log("准备启动自动化...")
        self.log(f"执行条数: {'全部' if limit == 0 else limit}")
        self.log("窗口将自动最小化，完成后恢复")
        self.log("安全提示: 将鼠标移到屏幕左上角可紧急停止")
        self.log("=" * 40)

        # 最小化窗口
        self.root.iconify()

        # 在新线程中运行
        def run_bot():
            try:
                import time
                time.sleep(1)  # 等待窗口最小化
                from main_bot import AutomationBot
                bot = AutomationBot(str(self.config_path))
                bot.run(limit=limit)
                self.log("运行完成!")
            except Exception as e:
                self.log(f"错误: {e}")
            finally:
                # 恢复窗口
                self.root.after(0, self.root.deiconify)

        threading.Thread(target=run_bot, daemon=True).start()

    def refresh_all(self):
        """刷新所有数据"""
        self.refresh_steps()
        self.check_ready()

        # 预览已有Excel
        excel_path = self.config.get('excel', {}).get('file_path', '')
        if excel_path:
            full_path = self.base_dir / excel_path
            if full_path.exists():
                self.preview_excel(full_path)

    def run(self):
        """运行主循环"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ControlPanel()
    app.run()
