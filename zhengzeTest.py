import tkinter as tk
import re
from tkinter import ttk, scrolledtext, messagebox


class RegexDebugger:
    def __init__(self, master):
        self.master = master
        master.title("正则表达式调试工具")
        self.create_widgets()
        self.setup_layout()

    def create_widgets(self):
        # 正则表达式输入
        self.regex_entry = ttk.Entry(self.master, width=60, font=('Courier', 10))
        self.regex_entry.bind('<KeyRelease>', self.update_matches)

        # 测试文本输入
        self.test_text = scrolledtext.ScrolledText(
            self.master, width=80, height=15,
            font=('Courier', 10), wrap=tk.WORD
        )
        self.test_text.bind('<KeyRelease>', self.update_matches)
        self.test_text.tag_configure('highlight', background='#FFFF00')

        # 结果展示
        self.result_tree = ttk.Treeview(self.master, columns=('start', 'end', 'match', 'groups'), show='headings',
                                        height=8)
        self.result_tree.heading('start', text='起始位置')
        self.result_tree.heading('end', text='结束位置')
        self.result_tree.heading('match', text='匹配内容')
        self.result_tree.heading('groups', text='分组信息')
        self.result_tree.column('start', width=80)
        self.result_tree.column('end', width=80)
        self.result_tree.column('match', width=350)
        self.result_tree.column('groups', width=500)

        # 错误信息
        self.error_label = ttk.Label(self.master, text='', foreground='red')

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.master, textvariable=self.status_var, anchor=tk.W)

    def setup_layout(self):
        # 布局管理
        ttk.Label(self.master, text="正则表达式:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.regex_entry.grid(row=1, column=0, sticky='ew', padx=5, pady=2)

        ttk.Label(self.master, text="测试文本:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.test_text.grid(row=3, column=0, sticky='nsew', padx=5, pady=2)

        ttk.Label(self.master, text="匹配结果:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
        self.result_tree.grid(row=5, column=0, sticky='nsew', padx=5, pady=2)

        self.error_label.grid(row=6, column=0, sticky='w', padx=5, pady=2)
        self.status_bar.grid(row=7, column=0, sticky='ew', padx=5, pady=2)

        # 配置网格权重
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(3, weight=1)
        self.master.grid_rowconfigure(5, weight=1)

    def update_matches(self, event=None):
        # 清除之前的匹配结果
        self.test_text.tag_remove('highlight', '1.0', 'end')
        self.result_tree.delete(*self.result_tree.get_children())
        self.error_label.config(text='')
        self.status_var.set('')

        # 获取输入内容
        regex_str = self.regex_entry.get()
        test_text = self.test_text.get('1.0', 'end-1c')

        if not regex_str or not test_text:
            return

        try:
            pattern = re.compile(regex_str)
        except re.error as e:
            self.error_label.config(text=f'正则表达式错误: {e}')
            return

        # 生成位置索引映射
        text_length = len(test_text)
        pos_to_index = []
        if text_length > 0:
            index = '1.0'
            pos_to_index.append(index)
            for _ in range(text_length - 1):
                index = self.test_text.index(f"{index}+1c")
                pos_to_index.append(index)

        # 执行匹配
        matches = []
        try:
            matches = list(pattern.finditer(test_text))
        except Exception as e:
            self.error_label.config(text=f'匹配错误: {e}')
            return

        # 处理匹配结果
        match_count = 0
        for match in matches:
            start, end = match.start(), match.end()
            if start >= text_length or end > text_length:
                continue

            # 计算高亮范围
            start_index = pos_to_index[start]
            end_index = self.test_text.index(f"{pos_to_index[end - 1]}+1c") if end > 0 else pos_to_index[0]

            # 添加高亮
            self.test_text.tag_add('highlight', start_index, end_index)

            # 添加结果到树形视图
            groups = match.groups() or ()
            self.result_tree.insert('', 'end', values=(
                start_index,
                self.test_text.index(end_index),
                repr(match.group())[1:-1],
                ', '.join(repr(g)[1:-1] for g in groups) if groups else '无'
            ))
            match_count += 1

        # 更新状态栏
        self.status_var.set(f'共找到 {match_count} 处匹配')
        if not matches:
            self.error_label.config(text='未找到匹配内容')


if __name__ == '__main__':
    root = tk.Tk()
    root.geometry('900x700')
    app = RegexDebugger(root)
    root.mainloop()