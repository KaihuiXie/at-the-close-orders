# -*- coding: utf-8 -*-
"""
Created on Tue Jul 27 11:15:30 2021

@author: Xie Kaihui
"""

import easyquotation
import pandas as pd
import re
import tkinter as tk
import tkintertable as tkt
import datetime
import eventlet


def auto_sieve():
    attemp = 0
    
    while attemp < 16:
        try:
            eventlet.monkey_patch()
            with eventlet.Timeout(5,False):
                quotation = easyquotation.use('tencent') # 新浪 ['sina'] 腾讯 ['tencent', 'qq']
                df = pd.DataFrame(quotation.all_market).T
                break
            raise RuntimeError()
        except:
            attemp += 1

    if attemp == 16:
        tk.messagebox.showinfo('提示', '数据载入超时，请检查网络连接！')
        return None

    global time
    time = datetime.datetime.now()

    df = df.dropna(axis = 1, how = 'all')

    #step 1
    df['涨跌(%)'] = df['涨跌(%)'].astype('float64')
    df = df[(df['涨跌(%)'] >= 3) & (df['涨跌(%)'] <= 5)]

    #Step 2
    df = df[df['量比'] >= 1]

    #Step 3
    df = df[(df['turnover'] >= 5) & (df['turnover'] <= 10)]

    #Step 4
    df['流通市值'] = df['流通市值'].astype('float64')
    df = df[(df['流通市值'] >= 50) & (df['流通市值'] <= 200)]

    #keep A-shares
    code_prefix = ['sz00', 'sz30', 'sh60', 'sh68']

    for code in df.index:
        if re.findall('(s[zh]{1}[0-9]{2}).',code)[0] not in code_prefix:
            df = df.drop(code)

    #delete prefixes
    for code in df.index:
        df = df.rename(index={code:re.findall('s[zh]{1}([0-9]{6})', code)[0]})

    return df

def manual_sieve():
    imp_fpath = tk.filedialog.askopenfilename()
    if imp_fpath == '':
        return None
    try:
        df = pd.read_table(imp_fpath, encoding='gbk', na_values='--  ')
    except:
        tk.messagebox.showinfo('提示', '找不到文件！')
        return None
    
    if len(df) < 4000:
        tk.messagebox.showinfo('提示','导入数据过少，请检查在导出数据时是否已选择“报表中所有数据（所有栏目）”')
        return None
    
    cols = ['代码', '涨幅%', '量比', '换手%', '流通市值']
    for col in cols:
        if col not in df.columns:
            tk.messagebox.showinfo('提示','导入数据过少，请检查在导出数据时是否已选择“报表中所有数据（所有栏目）”')
            return None
    
    global time
    time = datetime.datetime.now()
    
    df = df.dropna(axis = 1, how = 'all')
    df = df.set_index('代码')
    df = df.drop('数据来源:通达信')

    #Step 1
    df['涨幅%'] = df['涨幅%'].astype('float64')
    df = df[(df['涨幅%'] >= 3) & (df['涨幅%'] <= 5)]

    #Step 2
    df = df[df['量比'] >= 1]
    
    #Step 3
    df = df[(df['换手%'] >= 5) & (df['换手%'] <= 10)]
    
    #Step 4
    pattern = r'([0-9.-]+)亿'
    for i in df.index: 
        df.loc[i, '流通市值'] = re.findall(pattern, df.loc[i, '流通市值'])[0]
    df['流通市值'] = df['流通市值'].astype('float64')
    df = df[(df['流通市值'] >= 50) & (df['流通市值'] <= 200)]
    
    return df



#UI windows
def save_click():
    default_file_name = re.sub(':','-',str(time))
    fpath = tk.filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel','.xlsx')],initialfile=default_file_name)
    if fpath != '':
        try:
            df.to_excel(fpath)
            tk.messagebox.showinfo('提示','保存成功！')
        except:
            tk.messagebox.showinfo('提示', '保存失败：找不到路径！'+fpath)
            
def back_click():
    lbl_.configure(text='欢迎！')
    tframe.grid_forget()
    table.destroy()
    btn_save.grid_forget()
    btn_back.grid_forget()
    btn_main.grid(column=0,row=1)
    btn_manual.grid(column=1,row=1)

def main_click():
    lbl_.configure(text='正在载入...')
    btn_main.grid_forget()
    btn_manual.grid_forget()
    global df
    df = auto_sieve()
    if df is None:
        lbl_.configure(text='欢迎！')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
        return
    lbl_.configure(text='筛选结果')
    if len(df) == 0:
        tk.messagebox.showinfo('提示', '此时间段未找到符合筛选标准的股票，请稍后再试！')
        lbl_.configure(text='欢迎！')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
    else:
        global tframe
        tframe = tkt.Frame(window)
        tframe.grid(column=0,row=1)
        dic = df.T.to_dict(orient='dict')
        global table
        table = tkt.TableCanvas(tframe, data = dic)
        table.show()
        global btn_save
        btn_save = tk.Button(window, text='保存',command=save_click)
        btn_save.grid(column=0,row=15)
        global btn_back
        btn_back = tk.Button(window, text='返回',command=back_click)
        btn_back.grid(column=1,row=15)
        
def back_click_1():
    lbl_.configure(text='欢迎！')
    lbl_mn_1.grid_forget()
    lbl_mn_2.grid_forget()
    lbl_mn_3.grid_forget()
    btn_back_1.grid_forget()
    btn_choose.grid_forget()
    btn_main.grid(column=0,row=1)
    btn_manual.grid(column=1,row=1)
        
def choose_file():
    lbl_mn_1.grid_forget()
    lbl_mn_2.grid_forget()
    lbl_mn_3.grid_forget()
    btn_back_1.grid_forget()
    btn_choose.grid_forget()
    global df
    df = manual_sieve()
    if df is None:
        lbl_.configure(text='欢迎！')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
        return
    lbl_.configure(text='筛选结果')
    if len(df) == 0:
        tk.messagebox.showinfo('提示', '此时间段未找到符合筛选标准的股票，请稍后再试！')
        lbl_.configure(text='欢迎！')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
    else:
        global tframe
        tframe = tkt.Frame(window)
        tframe.grid(column=0,row=1)
        dic = df.T.to_dict(orient='dict')
        global table
        table = tkt.TableCanvas(tframe, data = dic)
        table.show()
        global btn_save
        btn_save = tk.Button(window, text='保存',command=save_click)
        btn_save.grid(column=0,row=15)
        global btn_back
        btn_back = tk.Button(window, text='返回',command=back_click)
        btn_back.grid(column=1,row=15)

def manual_click():
    lbl_.configure(text='手动筛选操作步骤：',font=26)
    lbl_.grid(sticky='w')
    global lbl_mn_1,lbl_mn_2,lbl_mn_3
    lbl_mn_1 = tk.Label(window, text='1. 在银河证券海王星界面中输入“60”，回车，再输入“34”，回车',font=26)
    lbl_mn_2 = tk.Label(window, text='2. 在弹出的对话框的“格式文本文件”中选择“报表中所有数据（所有栏目）”，\n点击“浏览”选择一个保存位置，确认后点击“导出”', font=26)
    lbl_mn_3 = tk.Label(window, text='3.点击下方“选择文件”将从银河证券海王星导出的文件导入', font=26)
    lbl_mn_1.grid(column=0,row=1,sticky='w')
    lbl_mn_2.grid(column=0,row=2,sticky='w')
    lbl_mn_3.grid(column=0,row=3,sticky='w')
    btn_main.grid_forget()
    btn_manual.grid_forget()
    
    global btn_choose,btn_back_1
    btn_choose = tk.Button(window, text='选择文件', command=choose_file)
    btn_choose.grid(column=0,row=4,sticky='w')
    btn_back_1 = tk.Button(window, text='返回', command=back_click_1)
    btn_back_1.grid(column=1,row=4,sticky='w')
    
window = tk.Tk(className='下午2：30行情筛选')
window.geometry('1020x480')

lbl_ = tk.Label(window, text = '欢迎！', font = 50)
lbl_.grid(column=0,row=0)
   

btn_main = tk.Button(window, text = '自动筛选', command = main_click)
btn_main.grid(column=0,row=1)
btn_manual = tk.Button(window, text = '手动筛选', command = manual_click)
btn_manual.grid(column=1,row=1)

window.mainloop()
