import logging
import tkinter as tk
from tkinter import scrolledtext


class TextHandler(logging.Handler):
    """自定义日志处理器，将日志输出到Tkinter文本框"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.text_widget.config(state=tk.DISABLED)

        # 设置日志格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.setFormatter(formatter)

    def emit(self, record):
        """处理日志记录"""
        msg = self.format(record)

        def append():
            """线程安全的文本追加操作"""
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)  # 滚动到底部
            self.text_widget.config(state=tk.DISABLED)

        # 确保在主线程更新UI
        if self.text_widget.winfo_exists():
            self.text_widget.after(0, append)