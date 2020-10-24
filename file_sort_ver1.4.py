# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog
#import codecs
import pandas as pd
#import numpy as np
import os
import re
import cv2
from tkinter import ttk 

root = tk.Tk()

file_path1 = ()


dict_alp_num = {"A":0,"B":1,"C":2,"D":3,"E":4,"F":5,"G":6,"H":7,"I":8,"J":9,
                "K":10,"L":11,"M":12,"N":13,"O":14,"P":15,"Q":16,"R":17,
                "S":18,"T":19,"U":20,"V":21,"W":22,"X":23,"Y":24,"Z":25}


def hukusuu(x):
    s = x.split(",")
    l = len(s)
    return s,l

def word_detect(x):
    a = len(x)
    print(a)
    if len(x) >= 3:
        lv4 = tk.Label(root,text="please enter again", fg="red")
        lv4.place(x=80,y=205,width=150,height=20)
        lv4.configure(background='gray')
        print(1)
    elif x.isalpha() == False:
        lv4 = tk.Label(root,text="please enter alphabet only", fg="red")
        lv4.place(x=80,y=205,width=150,height=20)
        lv4.configure(background='gray')
        print(2)
        
def word_trans(x):
    a = len(x)
    if a == 1:
        clmns = dict_alp_num[x] - 1
        return clmns
    else:    
        tran = list(x)
        clmns = ((dict_alp_num[tran[0]]+1) * 26) + dict_alp_num[tran[1]] - 1
        return clmns

    
def write(key,key2,file_path1,sheets):
    length_data = len(file_path1)
    df_final = pd.DataFrame([0])
    if sheets == "":
        sheets = None
    s = chk.get()
    print(type(s),s)
    s = int(s)
    print(type(s),s)
    exten = os.path.splitext(file_path1[0])
    print(exten)
    if entry4.get() == "":
        namesave = "test"
    else:
        namesave = entry4.get()
    if exten[1] == ".xlsx" or ".XLSX":
        aa = 2
        key2 = key2 + 1
        print('key2=',key2)
        for num in range(length_data):
            print(file_path1[num])
            #df = pd.read_csv(file_path1[num],header=s,index_col=0,engine="python")
            print(sheets)
            df = pd.read_excel(file_path1[num],header=s,index_col=None,sheet_name=sheets)
            print(df)
            num_file = re.sub("\\D", "", file_path1[num])
            print(num_file)
            print(df.iloc[:,[key2]].dropna(how='any'))
            df_wxs = df.iloc[:,[key2]].dropna(how='any')
            df_wxs.columns = [num_file+key]
            print(df_wxs)
            #df_final = df_final.join(df_wx)
            df_final = pd.concat([df_final,df_wxs],axis=1)
            print(df_final)
        #df_final.to_csv('C:/Users/naona/OneDrive/python data/test.csv')
        df_final = df_final.drop(df_final.columns[[0]], axis=1)
        print(df_final)
        save_ = os.path.join(entry5.get(),namesave+'.xlsx')
        print(save_)
        return save_,df_final,aa
    elif exten[1] == ".csv" or ".CSV":
        key2 = key2 + 1
        print('key2=',key2)
        aa = 1
        for num in range(length_data):
            print(file_path1[num])
            print("s",s)
            df = pd.read_csv(file_path1[num],header=s,error_bad_lines=False,engine="python")
            print(df)
            num_file = re.sub("\\D", "", file_path1[num])
            print(num_file)
            print(type(num_file))
            print(key)
            print(type(key))
            print(df.iloc[:,[key2]].dropna(how='any'))
            df_wxs = df.iloc[:,[key2]].dropna(how='any')
            #df_wxs.columns = ['crank']
            df_wxs.columns = [str(num_file+key)]
            print(df_wxs)
            #df_final = df_final.join(df_wx)
            df_final = pd.concat([df_final,df_wxs],axis=1)
            print(df_final)
        df_final = df_final.drop(df_final.columns[[0]], axis=1)
        print(df_final)
        save_ = os.path.join(entry5.get(),namesave+'.csv')
        print(save_)
        return save_,df_final,aa
    else:
        lv5 = tk.Label(root,text="file extension is not valued", fg="red")
        lv5.place(x=70,y=330,width=150,height=20)
        lv5.configure(background='gray')
        
def save(dirct,data_fre,aa):
    print(data_fre)
    if aa == 2:
        data_fre.to_excel(dirct)
    elif aa == 1:
        data_fre.to_csv(dirct)
        

def clicked1(x):
    global file_path1
    idir = 'C:/User/shiba/desktop'
    type=[('excelファイル','*.xlsx'),('csvファイル','*.csv')]
    file_path = filedialog.askopenfilenames(filetypes = type ,initialdir = idir)
    print((file_path))
    length_file = len(file_path)
    for num1 in range(length_file):
        entry1.insert(tk.END,file_path[num1])
        stoped_number = entry1.get()
        if len(stoped_number) > 500:
            break
    entry1.configure(state='readonly')
    file_path1 = file_path
        
def clicked2(event):
    entry1.delete(0,tk.END)
    entry1.configure(state='normal')
    results = chk.get()
    print(results)
        
def clicked3(event):
    print(file_path1)
    key = entry2.get().upper()
    key_huku,pass_key = hukusuu(key)
    if pass_key == 1:
        print("指定した列は一つです")
        sheets = entry3.get()
        word_detect(key)
        key2 = word_trans(key)
        save_file,df_final1,aaa = write(key_huku[0],key2,file_path1,sheets)
        print(df_final1)
        save(save_file,df_final1,aaa)
    elif pass_key > 1:
        print("指定した列は二つです")
        df_final2 = pd.DataFrame([0])
        for key_num in range(pass_key):
            sheets = entry3.get()
            word_detect(key_huku[key_num])
            key2 = word_trans(key_huku[key_num])
            save_file,df_final1,aaa = write(key_huku[key_num],key2,file_path1,sheets)
            df_final2 = pd.concat([df_final2,df_final1],axis=1)
            print("一つ目")
            print(df_final2)
        df_final2 = df_final2.drop(df_final2.columns[[0]], axis=1)
        print("2つ目")
        save(save_file,df_final2,aaa)
        
    
def clicked4(event):
    root.destroy()
    
def clicked5(event):
    global dir_save
    idir = 'C:/User/shiba/desktop'
    dir_save = filedialog.askdirectory(initialdir = idir)
    if dir_save != "":
        entry5.configure(state='normal')
        entry5.delete(0,tk.END)
        entry5.insert(tk.END,dir_save)
        entry5.configure(state='readonly')
    
def clicked6(event):         # Create window with freedom of dimensions
    im = cv2.imread("C:/Users/naona/OneDrive/python data/example.jpg")                     # Read image
    imS = cv2.resize(im, (600,600))                    # Resize image
    cv2.imshow("output", imS)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    
#ウィンドウの大きさの設定    
root.configure(background='gray')
root.title("Sort program")
root.geometry("500x450")


#テキストの設定
lv1 = tk.Label(root,text="Excel sort program")
lv1.place(x=350,y=20,width=150,height=30)
lv1.configure(background='gray')
lv2 = tk.Label(root,text="select columns:複数列欲しいときは、\nカンマ（,）で区切ってね。")
lv2.place(x=20,y=165,width=200,height=50)
lv2.configure(background='gray')
lv3 = tk.Label(root,text="sheet name:\ncsvファイルの時入力しなくてもいい。")
lv3.place(x=10,y=220,width=230,height=30)
lv3.configure(background='gray')
lv6 = tk.Label(root,text="←enter save file name")
lv6.place(x=300,y=280,width=150,height=20)
lv6.configure(background='gray')
lv7 = tk.Label(root,text="if not enter, file name will be 'test'")
lv7.place(x=260,y=295,width=250,height=20)
lv7.configure(background='gray')
lv8 = tk.Label(root,text="AT:totalQ AW:smooth_ROHR",font=(None, 10))
lv8.place(x=300,y=170,width=200,height=15)
lv8.configure(background='gray')
lv9 = tk.Label(root,text="AS:ROHR BZ:main pressure",font=(None, 10))
lv9.place(x=300,y=185,width=200,height=15)
lv9.configure(background='gray')
lv10 = tk.Label(root,text="BF:pilot pressure",font=(None, 10))
lv10.place(x=300,y=200,width=200,height=15)
lv10.configure(background='gray')
lv11 = tk.Label(root,text="calc,clac2,data...",font=(None, 10))
lv11.place(x=300,y=230,width=200,height=15)
lv11.configure(background='gray')
lv12 = tk.Label(root,text="入力例（画像はｘを押さずに適当なキーボードを一つ押したらきえます。）"
                ,font=(None, 8), fg="red")
lv12.place(x=130,y=57,width=400,height=15)
lv12.configure(background='gray')
lv13 = tk.Label(root,text="select first index")
lv13.place(x=20,y=20,width=150,height=30)
lv13.configure(background='gray')


#ファイルの表示画面を出す。
entry1 = tk.Entry(root)
entry2 = tk.Entry(root)
entry3 = tk.Entry(root)
entry4 = tk.Entry(root)
entry5 = tk.Entry(root)
entry1.place(x=30,y=80,width=250,height=80)
entry2.place(x=220,y=175,width=80,height=30)
entry3.place(x=220,y=220,width=80,height=30)
entry4.place(x=30,y=280,width=250,height=30)
entry5.place(x=30,y=340,width=250,height=30)
#entry.configure(state='readonly')

#ボタンの設定
botton1 = tk.Button(root,text="select files")
botton2 = tk.Button(root,text="delete list")
botton3 = tk.Button(root,text="make excel")
botton4 = tk.Button(root,text="cancel")
botton5 = tk.Button(root,text="select save directry")
botton6 = tk.Button(root,text="example")

#ボタンのイベント情報
botton1.bind("<Button-1>",clicked1)
botton2.bind("<Button-1>",clicked2)
botton3.bind("<Button-1>",clicked3)
botton4.bind("<Button-1>",clicked4)
botton5.bind("<Button-1>",clicked5)
botton6.bind("<Button-1>",clicked6)

#ボタンの位置
botton1.place(x=300,y=85,width=150,height=30)
botton2.place(x=300,y=125,width=150,height=30)
botton3.place(x=70,y=400,width=150,height=30)
botton4.place(x=290,y=400,width=150,height=30)
botton5.place(x=300,y=340,width=150,height=30)
botton6.place(x=100,y=55,width=60,height=20)

#comboboxの設定
var = tk.IntVar()
test = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30]
chk = ttk.Combobox(root,values=test)
chk.place(x=150, y=20)


#初期値の設定
desktop_path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\Desktop"
entry5.insert(tk.END,desktop_path)
entry5.configure(state='readonly')


root.mainloop()



    

    
