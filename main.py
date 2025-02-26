import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import akshare as ak
import pandas as pd
from datetime import datetime
import os
from openpyxl import Workbook
from threading import Thread

class StockAnalyzer(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title('股票数据分析工具')
        self.geometry('800x600')
        
        # 设置主题样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 自定义样式
        style.configure('TLabel', padding=5, font=('微软雅黑', 10))
        style.configure('TEntry', padding=5)
        style.configure('TButton', padding=5, font=('微软雅黑', 10))
        style.configure('Custom.TFrame', background='#f0f0f0')
        
        # 创建主框架
        self.main_frame = ttk.Frame(self, style='Custom.TFrame')
        self.main_frame.pack(padx=20, pady=20, fill='both', expand=True)
        
        # 创建输入区域
        self.create_input_area()
        
        # 创建日志区域
        self.create_log_area()
        
        # 创建进度条
        self.progress = ttk.Progressbar(self.main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky='ew', padx=5, pady=5)
        
    def create_input_area(self):
        # 创建输入框架
        input_frame = ttk.LabelFrame(self.main_frame, text='数据查询', padding=10)
        input_frame.grid(row=0, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        
        # 股票代码输入
        ttk.Label(input_frame, text='股票代码：').grid(row=0, column=0, sticky='w')
        self.stock_code = ttk.Entry(input_frame, width=15)
        self.stock_code.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # 日期选择
        ttk.Label(input_frame, text='起始日期：').grid(row=0, column=2, sticky='w')
        self.start_date = DateEntry(input_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='y-mm-dd')
        self.start_date.grid(row=0, column=3, padx=5, pady=5)
        self.start_date.set_date('2023-01-01')
        
        ttk.Label(input_frame, text='结束日期：').grid(row=0, column=4, sticky='w')
        self.end_date = DateEntry(input_frame, width=12, background='darkblue',
                                foreground='white', borderwidth=2, date_pattern='y-mm-dd')
        self.end_date.grid(row=0, column=5, padx=5, pady=5)
        
        # 按钮
        self.fetch_btn = ttk.Button(input_frame, text='获取数据', command=self.start_fetch)
        self.fetch_btn.grid(row=0, column=6, padx=10, pady=5)
        
    def create_log_area(self):
        # 创建日志框架
        log_frame = ttk.LabelFrame(self.main_frame, text='运行日志', padding=10)
        log_frame.grid(row=4, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        
        # 创建文本框和滚动条
        self.log_text = tk.Text(log_frame, height=15, width=80, wrap=tk.WORD,
                              font=('Consolas', 9), bg='#f8f8f8')
        scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 放置文本框和滚动条
        self.log_text.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # 配置grid权重
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
    def log(self, message):
        self.log_text.insert('end', f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - {message}\n')
        self.log_text.see('end')
        self.update_idletasks()
        
    def start_fetch(self):
        self.fetch_btn['state'] = 'disabled'
        self.progress.start(10)
        Thread(target=self.fetch_data).start()
        
    def fetch_data(self):
        try:
            stock_code = self.stock_code.get().strip()
            start_date = self.start_date.get_date().strftime('%Y-%m-%d')
            end_date = self.end_date.get_date().strftime('%Y-%m-%d')
            
            self.log(f'输入参数检查 - 股票代码: {stock_code}, 起始日期: {start_date}, 结束日期: {end_date}')
            
            if not stock_code:
                self.log('请输入股票代码')
                return
            
            if start_date > end_date:
                self.log('起始日期不能晚于结束日期')
                messagebox.showerror('错误', '起始日期不能晚于结束日期')
                return
            
            try:
                formatted_code = self.format_stock_code(stock_code)
                self.log(f'开始获取股票 {formatted_code} 的数据...')
                
                try:
                    df = ak.stock_zh_a_hist(symbol=formatted_code, start_date=start_date, end_date=end_date)
                    self.log(f'成功调用akshare接口获取数据')
                except Exception as e:
                    self.log(f'调用akshare接口失败: {str(e)}')
                    raise ValueError(f'获取股票数据失败: {str(e)}')
                
                if df.empty:
                    self.log('未获取到数据')
                    return
                
                self.log('数据获取成功，开始计算技术指标...')
                df = self.calculate_indicators(df)
                
                self.log('开始导出Excel文件...')
                filename = self.export_to_excel(df, formatted_code)
                
                self.log(f'数据已成功导出到文件：{filename}')
                messagebox.showinfo('成功', f'数据已导出到文件：{filename}')
                
            except ValueError as ve:
                self.log(f'格式化或获取数据时发生错误：{str(ve)}')
                messagebox.showerror('错误', str(ve))
            
        except Exception as e:
            self.log(f'发生未知错误：{str(e)}')
            messagebox.showerror('错误', f'发生未知错误：{str(e)}')
        
        finally:
            self.fetch_btn['state'] = 'normal'
            self.progress.stop()
            self.progress['value'] = 0

    def format_stock_code(self, code):
        self.log(f'开始格式化股票代码: {code}')
        try:
            code = str(code).zfill(6)
            self.log(f'补零后的股票代码: {code}')
            
            if not code.isdigit():
                raise ValueError(f'股票代码必须为数字，当前输入: {code}')
                
            # 直接返回6位数字代码
            if code.startswith('6') or code.startswith(('0', '3')):
                return code
            else:
                raise ValueError(f'不支持的股票代码格式: {code}，股票代码必须以0、3或6开头')
            
        except Exception as e:
            self.log(f'股票代码格式化失败: {str(e)}')
            raise

    def calculate_indicators(self, df):
        # 计算MA5、MA10、MA20
        df['MA5'] = df['收盘'].rolling(window=5).mean()
        df['MA10'] = df['收盘'].rolling(window=10).mean()
        df['MA20'] = df['收盘'].rolling(window=20).mean()
        
        # 计算MACD
        exp1 = df['收盘'].ewm(span=12, adjust=False).mean()
        exp2 = df['收盘'].ewm(span=26, adjust=False).mean()
        df['MACD'] = exp1 - exp2
        df['Signal'] = df['MACD'].ewm(span=9, adjust=False).mean()
        df['Histogram'] = df['MACD'] - df['Signal']
        
        return df
    
    def export_to_excel(self, df, stock_code):
        # 创建保存目录
        if not os.path.exists('output'):
            os.makedirs('output')
            
        try:
            # 获取股票名称
            stock_info = ak.stock_info_a_code_name()
            stock_name = stock_info[stock_info['代码'] == stock_code]['名称'].values[0]
            
            # 生成文件名（使用查询的结束日期）
            query_date = self.end_date.get_date().strftime('%Y%m%d')
            filename = f'output/{stock_code}_{stock_name}_{query_date}.xlsx'
            
            # 保存到Excel（覆盖已存在的文件）
            df.to_excel(filename, index=True)
            
            return filename
            
        except Exception as e:
            self.log(f'获取股票名称失败: {str(e)}')
            # 如果获取名称失败，使用默认的文件名格式
            filename = f'output/{stock_code}_{self.end_date.get_date().strftime("%Y%m%d")}.xlsx'
            df.to_excel(filename, index=True)
            return filename

if __name__ == '__main__':
    app = StockAnalyzer()
    app.mainloop()