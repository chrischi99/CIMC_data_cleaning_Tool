from __future__ import print_function
import pysolr
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

root= tk.Tk()

# Set the main canvas for ui
canvas1 = tk.Canvas(root, width = 500, height = 500, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

# Title and background color
label1 = tk.Label(root, text='模糊搜索工具', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(250, 120, window=label1)

#User input a Excel file
def getExcel ():
    global read_file
    global import_file_path
    import_file_path = filedialog.askopenfilename()
    try:
        read_file = pd.read_excel(import_file_path)
        messagebox.showinfo("文件上传", "文件上传成功")
    except NameError:
        messagebox.showinfo("文件上传", "文件上传失败")
        
    

# Upload button
browseButton_Excel = tk.Button(text="      上传Excel文件     ", command=getExcel, bg='green', fg='black', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 180, window=browseButton_Excel)

# User input a csv file
# def getCSV():
#     global read_file
#     global import_file_path
#     import_file_path = filedialog.askopenfilename()
#     try:
#         read_file = pd.read_csv(import_file_path)
#         messagebox.showinfo("文件上传", "文件上传成功")
#     except NameError:
#         messagebox.showinfo("文件上传", "文件上传失败")

# # Upload button
# browseButton_CSV = tk.Button(text="      上传CSV文件     ", command=getCSV, bg='green', fg='black', font=('helvetica', 12, 'bold'))
# canvas1.create_window(250, 250, window=browseButton_CSV)

# Fuzzy Search
def convertToExcel():
    global read_file
    global import_file_path
    try:
        read_file
    except NameError:
        messagebox.showinfo("未发现文件", "未发现文件!请上传文件")

    id_name = simpledialog.askstring(title = "柱子名称" , prompt = "请输入需要模糊查询的柱子名称")

    result_list = []
    # Setup the Solr instance from the database 
    solr = pysolr.Solr('http://localhost:8983/solr/CIMC/', timeout=10)

    # Do a health check.
    solr.ping()
    # Iterate through the csv file
    for i in read_file[id_name]:
        # Preparation for exact match
        company_name = 'name:"'+i+'"'
        #Exact match search
        results = solr.search(company_name)
        exact_search_num = len(results)
        if (exact_search_num == 1):
            for result in results:
                # Edge cases
                if ("龙口中集" in result[id_name]):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result[id_name])
    
        else:
            # Perform a Fuzzy search on the database
            company_name_2 = "name:"+ i
            results = solr.search(company_name_2)
            count = 0
            for result in results:
                if count < 1:
                    # Edge cases
                    if ("龙口" in i):
                        result_list.append("Longkou CIMC Raffles Offshore Ltd")
                    else:
                        result_list.append(result[id_name])
                count += 1

    #Export the results to a excel file
    col_index = read_file.columns.get_loc(id_name)
    read_file.insert(col_index+1,"模糊搜索结果", result_list)
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=True)
    messagebox.showinfo("模糊搜索", "模糊搜索成功")

# Fuzzy Search button
saveAsButton_Excel = tk.Button(text='模糊搜索（新文件）', command=convertToExcel, bg='green', fg='black', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 250, window=saveAsButton_Excel)

def convertToExcel_2():
    global read_file
    global import_file_path
    try:
        read_file
    except NameError:
        messagebox.showinfo("未发现文件", "未发现文件!请上传文件")

    id_name = simpledialog.askstring(title = "柱子名称" , prompt = "请输入需要模糊查询的柱子名称")

    result_list = []
    # Setup the Solr instance from the database 
    solr = pysolr.Solr('http://localhost:8983/solr/CIMC/', timeout=10)

    # Do a health check.
    solr.ping()
    # Iterate through the csv file
    for i in read_file[id_name]:
        # Preparation for exact match
        company_name = 'name:"'+i+'"'
        #Exact match search
        results = solr.search(company_name)
        exact_search_num = len(results)
        if (exact_search_num == 1):
            for result in results:
                # Edge cases
                if ("龙口中集" in result[id_name]):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result[id_name])
    
        else:
            # Perform a Fuzzy search on the database
            company_name_2 = "name:"+ i
            results = solr.search(company_name_2)
            count = 0
            for result in results:
                if count < 1:
                    # Edge cases
                    if ("龙口" in i):
                        result_list.append("Longkou CIMC Raffles Offshore Ltd")
                    else:
                        result_list.append(result[id_name])
                count += 1

   #Export the results to a excel file
    col_index = read_file.columns.get_loc(id_name)
    read_file.insert(col_index+1,"模糊搜索结果", result_list)
    read_file.to_excel(import_file_path, index = None, header=True)
    messagebox.showinfo("模糊搜索", "模糊搜索成功")

saveAsButton_Excel_2 = tk.Button(text='模糊搜索', command=convertToExcel_2, bg='green', fg='black', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 320, window=saveAsButton_Excel_2)

# Exit application function
def exitApplication():
    MsgBox = tk.messagebox.askquestion ('你正在退出中','确定要退出吗',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()

# Exit button
exitButton = tk.Button (root, text='       退出程序     ',command=exitApplication, bg='brown', fg='red', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 390, window=exitButton)

# Repeat the mainloop to prevent exit
root.mainloop()