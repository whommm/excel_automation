# -*- coding: utf-8 -*-
"""
步骤配置工具 - 可视化界面
运行: python config_editor.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import yaml


class ConfigEditor:
    """配置编辑器"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("库存自动化 - 步骤配置器")
        self.root.geometry("800x600")

        self.config_path = Path("config.yaml")
        self.config = self.load_config()
        self.steps = self.config.get('steps', [])

        self.setup_ui()

    def load_config(self):
        """加载配置"""
        if self.config_path.exists():
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        return {'steps': []}

    def save_config(self):
        """保存配置"""
        self.config['steps'] = self.steps
        with open(self.config_path, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False)
        messagebox.showinfo("成功", "配置已保存!")

    def setup_ui(self):
        """构建界面"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill='x', padx=5, pady=5)

        ttk.Button(toolbar, text="添加步骤", command=self.add_step).pack(side='left', padx=2)
        ttk.Button(toolbar, text="删除选中", command=self.delete_step).pack(side='left', padx=2)
        ttk.Button(toolbar, text="上移", command=self.move_up).pack(side='left', padx=2)
        ttk.Button(toolbar, text="下移", command=self.move_down).pack(side='left', padx=2)
        ttk.Button(toolbar, text="保存配置", command=self.save_config).pack(side='right', padx=2)

        # 步骤列表
        list_frame = ttk.LabelFrame(self.root, text="操作步骤列表")
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ('序号', '名称', '动作', '目标')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        self.tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.tree.bind('<Double-1>', self.edit_step)

        self.refresh_list()

    def refresh_list(self):
        """刷新步骤列表"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, step in enumerate(self.steps, 1):
            self.tree.insert('', 'end', values=(
                i,
                step.get('name', ''),
                step.get('action', ''),
                step.get('target', step.get('text', ''))
            ))

    def get_selected_index(self):
        """获取选中项索引"""
        selection = self.tree.selection()
        if not selection:
            return None
        item = self.tree.item(selection[0])
        return int(item['values'][0]) - 1

    def delete_step(self):
        """删除步骤"""
        idx = self.get_selected_index()
        if idx is not None:
            del self.steps[idx]
            self.refresh_list()

    def move_up(self):
        """上移步骤"""
        idx = self.get_selected_index()
        if idx and idx > 0:
            self.steps[idx], self.steps[idx-1] = self.steps[idx-1], self.steps[idx]
            self.refresh_list()

    def move_down(self):
        """下移步骤"""
        idx = self.get_selected_index()
        if idx is not None and idx < len(self.steps) - 1:
            self.steps[idx], self.steps[idx+1] = self.steps[idx+1], self.steps[idx]
            self.refresh_list()

    def add_step(self):
        """添加新步骤"""
        self.open_step_dialog()

    def edit_step(self, event=None):
        """编辑步骤"""
        idx = self.get_selected_index()
        if idx is not None:
            self.open_step_dialog(idx)

    def open_step_dialog(self, edit_idx=None):
        """打开步骤编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑步骤" if edit_idx is not None else "添加步骤")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()

        # 获取现有数据
        step_data = self.steps[edit_idx] if edit_idx is not None else {}

        # 名称
        ttk.Label(dialog, text="步骤名称:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        name_var = tk.StringVar(value=step_data.get('name', ''))
        ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=0, column=1, padx=10, pady=5)

        # 动作类型
        ttk.Label(dialog, text="动作类型:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        action_var = tk.StringVar(value=step_data.get('action', 'click'))
        actions = ['click', 'double_click', 'type_text', 'press_key', 'wait', 'wait_for', 'clear_input']
        action_combo = ttk.Combobox(dialog, textvariable=action_var, values=actions, width=27)
        action_combo.grid(row=1, column=1, padx=10, pady=5)

        # 目标图片
        ttk.Label(dialog, text="目标图片:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        target_var = tk.StringVar(value=step_data.get('target', ''))
        target_frame = ttk.Frame(dialog)
        target_frame.grid(row=2, column=1, padx=10, pady=5, sticky='w')
        ttk.Entry(target_frame, textvariable=target_var, width=20).pack(side='left')

        def browse_image():
            path = filedialog.askopenfilename(
                initialdir="assets",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )
            if path:
                target_var.set(Path(path).name)

        ttk.Button(target_frame, text="浏览", command=browse_image).pack(side='left', padx=5)

        # 输入文本
        ttk.Label(dialog, text="输入文本:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
        text_var = tk.StringVar(value=step_data.get('text', ''))
        ttk.Entry(dialog, textvariable=text_var, width=30).grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="(可用 {code} 和 {quantity})").grid(row=3, column=2, padx=5, sticky='w')

        # 按键
        ttk.Label(dialog, text="按键:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
        key_var = tk.StringVar(value=step_data.get('key', ''))
        ttk.Entry(dialog, textvariable=key_var, width=30).grid(row=4, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="(如 enter, tab, ctrl+a)").grid(row=4, column=2, padx=5, sticky='w')

        # 等待时间
        ttk.Label(dialog, text="操作后等待(秒):").grid(row=5, column=0, padx=10, pady=5, sticky='e')
        wait_var = tk.StringVar(value=str(step_data.get('wait_after', 0.5)))
        ttk.Entry(dialog, textvariable=wait_var, width=30).grid(row=5, column=1, padx=10, pady=5)

        # 清空选项
        clear_var = tk.BooleanVar(value=step_data.get('clear_first', False))
        ttk.Checkbutton(dialog, text="输入前清空", variable=clear_var).grid(row=6, column=1, padx=10, pady=5, sticky='w')

        def save_step():
            step = {'name': name_var.get(), 'action': action_var.get()}

            if target_var.get():
                step['target'] = target_var.get()
            if text_var.get():
                step['text'] = text_var.get()
            if key_var.get():
                step['key'] = key_var.get()
            if clear_var.get():
                step['clear_first'] = True

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

            self.refresh_list()
            dialog.destroy()

        ttk.Button(dialog, text="保存", command=save_step).grid(row=7, column=1, pady=20)

    def run(self):
        """运行主循环"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ConfigEditor()
    app.run()
