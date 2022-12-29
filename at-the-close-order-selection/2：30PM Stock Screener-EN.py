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
from stopit import threading_timeoutable as timeoutable

@timeoutable(5)
def retrieve():
    quotation = easyquotation.use('tencent')  # 新浪 ['sina'] 腾讯 ['tencent', 'qq']
    return pd.DataFrame(quotation.all_market).T

def auto_sieve():
    attemp = 0
    while attemp < 16:
        try:
            df = retrieve()
            break
        except:
            attemp += 1

    if attemp == 16:
        tk.messagebox.showinfo('Alert', 'Timeout, please check out the Internet connection!')
        return None

    global time
    time = datetime.datetime.now()

    df = df.dropna(axis = 1, how = 'all')

    #step 1
    df['涨跌(%)'] = df['涨跌(%)'].astype('float64')
    df = df[(df['涨跌(%)'] >= 3) & (df['涨跌(%)'] <= 5)]  #['Percent Increase']

    #Step 2
    df = df[df['量比'] >= 1]  #['Quantity Relative Ratio']

    #Step 3
    df = df[(df['turnover'] >= 5) & (df['turnover'] <= 10)]

    #Step 4
    df['流通市值'] = df['流通市值'].astype('float64')
    df = df[(df['流通市值'] >= 50) & (df['流通市值'] <= 200)]  #['Market Capitalization']

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
        tk.messagebox.showinfo('Alert', 'File not found.')
        return None
    
    if len(df) < 4000:
        tk.messagebox.showinfo('Alert','Not enough data imported, please make sure "报表中所有数据（所有栏目）(all data)" is selected.')
        return None
    
    cols = ['代码', '涨幅%', '量比', '换手%', '流通市值']  #['Code', 'Percent Increase', 'Quantity Relative Ratio', 'Turnover Rate', 'Market Capitalization'
    for col in cols:
        if col not in df.columns:
            tk.messagebox.showinfo('Alert','Not enough data imported, please make sure "报表中所有数据（所有栏目）(all data)" is selected.')
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
            tk.messagebox.showinfo('Alert','Success!')
        except:
            tk.messagebox.showinfo('Alert', 'Failed：file path not exist: '+fpath)
            
def back_click():
    lbl_.configure(text='Welcome!')
    tframe.grid_forget()
    table.destroy()
    btn_save.grid_forget()
    btn_back.grid_forget()
    btn_main.grid(column=0,row=1)
    btn_manual.grid(column=1,row=1)

def main_click():
    lbl_.configure(text='Loading...')
    btn_main.grid_forget()
    btn_manual.grid_forget()
    global df
    df = auto_sieve()
    if df is None:
        lbl_.configure(text='Welcome!')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
        return
    lbl_.configure(text='Results')
    if len(df) == 0:
        tk.messagebox.showinfo('Alert', 'No stock found. Please try again later!')
        lbl_.configure(text='Welcome!')
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
        btn_save = tk.Button(window, text='Save',command=save_click)
        btn_save.grid(column=0,row=15)
        global btn_back
        btn_back = tk.Button(window, text='Back',command=back_click)
        btn_back.grid(column=1,row=15)
        
def back_click_1():
    lbl_.configure(text='Welcome!')
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
        lbl_.configure(text='Welcome！')
        btn_main.grid(column=0,row=1)
        btn_manual.grid(column=1,row=1)
        return
    lbl_.configure(text='Results')
    if len(df) == 0:
        tk.messagebox.showinfo('Alert', 'No stock found. Please try again later!')
        lbl_.configure(text='Welcome！')
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
        btn_save = tk.Button(window, text='Save',command=save_click)
        btn_save.grid(column=0,row=15)
        global btn_back
        btn_back = tk.Button(window, text='Back',command=back_click)
        btn_back.grid(column=1,row=15)

def manual_click():
    lbl_.configure(text='Manual screener steps: ',font=26)
    lbl_.grid(sticky='w')
    global lbl_mn_1,lbl_mn_2,lbl_mn_3
    lbl_mn_1 = tk.Label(window, text='1. Input "60" on the main page in the APP, press enter key, then input "34", press enter key',font=26)
    lbl_mn_2 = tk.Label(window, text='2. Choose "报表中所有数据（所有栏目）(all data)" in the prompt window "格式文本文件(file format)", \nclick"浏览(browse)" and choose a file path, then press, "导出(export)"', font=26)
    lbl_mn_3 = tk.Label(window, text='3. Click "Choose file" below and import the exported data', font=26)
    lbl_mn_1.grid(column=0,row=1,sticky='w')
    lbl_mn_2.grid(column=0,row=2,sticky='w')
    lbl_mn_3.grid(column=0,row=3,sticky='w')
    btn_main.grid_forget()
    btn_manual.grid_forget()
    
    global btn_choose,btn_back_1
    btn_choose = tk.Button(window, text='Choose file', command=choose_file)
    btn_choose.grid(column=0,row=4,sticky='w')
    btn_back_1 = tk.Button(window, text='Back', command=back_click_1)
    btn_back_1.grid(column=1,row=4,sticky='w')
    
window = tk.Tk(className='2：30PM Stock Screener')
window.geometry('1020x480')

lbl_ = tk.Label(window, text = 'Welcome!', font = 50)
lbl_.grid(column=0,row=0)
   

btn_main = tk.Button(window, text = 'Auto', command = main_click)
btn_main.grid(column=0,row=1)
btn_manual = tk.Button(window, text = 'Manual', command = manual_click)
btn_manual.grid(column=1,row=1)

window.mainloop()
