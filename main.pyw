import tkinter as tk
from tkinter import filedialog, dialog
import os
import numpy as np
import pandas as pd

class myWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('格式转换') # 标题
        self.window.geometry('500x500') # 窗口尺寸

        self.sb = tk.Scrollbar(self.window)
        self.sb.pack(side="right", fill="y")
        self.list_box = tk.Listbox(self.window, width=75, height=20, yscrollcommand=self.sb.set)
        self.list_box.pack()
        self.file_list = []

        bt1 = tk.Button(self.window, text='添加文件', width=15, height=2, command=self.add_file)
        bt1.pack(side=tk.LEFT, padx=20)
        bt2 = tk.Button(self.window, text='删除文件', width=15, height=2, command=self.delete_file)
        bt2.pack(side=tk.LEFT, padx=20)
        bt2 = tk.Button(self.window, text='导出文件', width=15, height=2, command=self.save_file)
        bt2.pack(side=tk.LEFT, padx=20)

        self.product_name_cmp_num = 30

    def add_file(self):
        file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('H:/')))
        if len(file_path) == 0:
            dialog.Dialog(None, {'title': '提示', 'text': '没有选中任何文件', 'bitmap': 'warning', 'default': 0,
            'strings': ('OK', 'Cancle')})
        elif file_path[-3:] != 'csv':
            dialog.Dialog(None, {'title': '提示', 'text': '请添加csv格式文件', 'bitmap': 'warning', 'default': 0,
            'strings': ('OK', 'Cancle')})
        elif file_path in self.file_list:
            dialog.Dialog(None, {'title': '提示', 'text': '文件已被添加', 'bitmap': 'warning', 'default': 0,
            'strings': ('OK', 'Cancle')})
        else:
            self.file_list.append(file_path)
            self.list_box.insert(tk.END, file_path)

    def delete_file(self):
        select_idx = self.list_box.curselection()
        if not select_idx or len(select_idx) == 0: return
        self.file_list.pop(select_idx[0])
        self.list_box.delete(select_idx)


    def pd_read_csv(self, file_path, rows=None, skiprow=None):
        for decode in ('gb18030', 'gbk','utf-8'):
            try:    
                data = pd.read_csv(file_path,encoding=decode,error_bad_lines=False, nrows=rows, skiprows=skiprow)
                return data
            except:
                pass
        print('%sencoding errow'.format(file_path))
        os._exit(0)

    def name_filter(self, df):
        df_name = df['Product Name '].str.slice(stop=self.product_name_cmp_num)
        df_com_l = np.empty(df_name.shape[0]+1, dtype=df_name.dtype)
        df_com_r = np.empty(df_name.shape[0]+1, dtype=df_name.dtype)
        df_com_l[:-1] = df_name
        df_com_r[1:] = df_name
        same_idx = df_com_l != df_com_r
        res = df[same_idx[:-1]].reset_index(drop=True)
        return res

    def save_file(self):
        result = None
        asin_set = []
        name_set = []
        if len(self.file_list) == 0: return
        first_two_lines = self.pd_read_csv(self.file_list[0], rows=1)
        for file in self.file_list:
            df_raw = self.pd_read_csv(file, skiprow=2)
            df = self.name_filter(df_raw)
            if 'ASIN ' not in df.columns:
                dialog.Dialog(None, {'title': '提示', 'text': file + '不包含列ASIN', 'bitmap': 'warning', 'default': 0,
            'strings': ('OK', 'Cancle')})
                os._exit(0)

            if 'Status ' not in df.columns:
                df['Status '] = 0

            if 'Note ' not in df.columns:
                df['Note '] = ''

            if 'Amazon Link' not in df.columns:
                df['Amazon Link '] = 'https://www.amazon.com/gp/product/' + df['ASIN ']

            if 'Dimensions ' not in df.columns: df['Dimensions '] = '0.00 x 0.00" x 0.00"" '
            NA_idx = np.where(df['Dimensions '] == 'N.A. ')
            df['Dimensions '][NA_idx[0]] = '0.00 x 0.00" x 0.00"" '
            df[['Length ', 'Width ', 'Height ']] = df['Dimensions '].str.split('x', expand=True)
            df['Width '] = df['Width '].astype(str).str[:-2]
            df['Height '] = df['Height '].astype(str).str[:-3]
            df['Volume '] = df['Length '].astype(float) * df['Width '].astype(float) * df['Height '].astype(float)

            if len(asin_set) == 0:
                result = df
                asin_set += df['ASIN '].tolist()
                name_set += df['Product Name '].str.slice(stop=self.product_name_cmp_num).tolist()
            else:
                asin_select_idx = np.isin(df['ASIN '], asin_set, invert=True)
                name_select_idx = np.isin(df['Product Name '].str.slice(stop=self.product_name_cmp_num), name_set, invert=True)
                select_idx = np.logical_and(asin_select_idx, name_select_idx)
                result = pd.concat([result, df[select_idx]])
                asin_set += df['ASIN '][select_idx].tolist()

        file_path = filedialog.asksaveasfilename(title=u'保存文件')
        if file_path[-3:] != 'csv':
            dialog.Dialog(None, {'title': '提示', 'text': '请保存为csv格式文件', 'bitmap': 'warning', 'default': 0,
            'strings': ('OK', 'Cancle')})
        elif file_path[-3:] == 'csv':
            first_two_lines.to_csv(file_path, index=None)
            result.to_csv(file_path, index=None, mode='a')
            os._exit(0)

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    mywindow = myWindow()
    mywindow.run()