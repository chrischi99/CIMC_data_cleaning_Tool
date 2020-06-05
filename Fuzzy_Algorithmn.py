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

# Fuzzy Search
def convertToExcel():
    global read_file
    global import_file_path
    no_company = 0
    try:
        read_file
    except NameError:
        messagebox.showinfo("未发现文件", "未发现文件!请上传文件")
    
    print("number of index of the file")
    print(len(read_file))

        
    id_name = simpledialog.askstring(title = "柱子名称" , prompt = "请输入需要模糊查询的柱子名称")

    result_list = []
    # Setup the Solr instance from the database 
    solr = pysolr.Solr('http://localhost:8983/solr/CIMC/', timeout=10)

    # Do a health check.
    solr.ping()
    print(read_file)
    # Iterate through the csv file
    for i in read_file[id_name]:
        # Preparation for exact match
        Abbreviation_name = 'Abbreviation:"'+i+'"'
        company_name = 'Name:"'+i+'"'
        #Exact match search
        Abbreviation_results = solr.search(Abbreviation_name)
        company_results = solr.search(company_name)
        exact_search_num = len(company_results)
        exact_search_num_2 = len(Abbreviation_results)
        if (exact_search_num == 1):
            for result in company_results:
                # Edge cases
                if ("龙口中集" in i):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result["Name"])
        elif(exact_search_num_2 == 1):
            for result in Abbreviation_results:
                # Edge cases
                if ("龙口中集" in i):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result["Name"])
        else:
            # Perform a Fuzzy search on the database
            Abbreviation_name_2 = "Abbreviation:"+ i
            company_name_2 = "Name:"+ i
            results = solr.search(company_name_2)
            count = 0
            if (len(results) == 0):
                result_list.append("无此公司")
            for result in results:
                if count < 1:
                    # Edge cases
                    if ("龙口" in i):
                        result_list.append("Longkou CIMC Raffles Offshore Ltd")
                    elif("烟台海工" in i):
                        result_list.append("中集海洋工程研究院有限公司")
                    elif("4s"in i):
                        result_list.append("无此公司")
                    else:
                        result_list.append(result["Name"])
                count += 1
    print("result of the list")
    print(len(result_list))
    print("number of weird company")
    print(no_company)

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
    import_file_path = filedialog.askopenfilename()
    try:
        read_file = pd.read_excel(import_file_path)
        messagebox.showinfo("文件上传", "文件上传成功")
    except NameError:
        messagebox.showinfo("文件上传", "文件上传失败")
        
    

# Upload button
browseButton_Excel = tk.Button(text="      上传Excel文件     ", command=getExcel, bg='green', fg='black', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 180, window=browseButton_Excel)

# Fuzzy Search
def convertToExcel_2():
    global read_file
    global import_file_path
    no_company = 0
    try:
        read_file
    except NameError:
        messagebox.showinfo("未发现文件", "未发现文件!请上传文件")
    
    print("number of index of the file")
    print(len(read_file))

        
    id_name = simpledialog.askstring(title = "柱子名称" , prompt = "请输入需要模糊查询的柱子名称")

    result_list = []
    # Setup the Solr instance from the database 
    solr = pysolr.Solr('http://localhost:8983/solr/CIMC/', timeout=10)

    # Do a health check.
    solr.ping()
    print(read_file)
    # Iterate through the csv file
    for i in read_file[id_name]:
        # Preparation for exact match
        Abbreviation_name = 'Abbreviation:"'+i+'"'
        company_name = 'Name:"'+i+'"'
        #Exact match search
        Abbreviation_results = solr.search(Abbreviation_name)
        company_results = solr.search(company_name)
        exact_search_num = len(company_results)
        exact_search_num_2 = len(Abbreviation_results)
        if (exact_search_num == 1):
            for result in company_results:
                # Edge cases
                if ("龙口中集" in i):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result["Name"])
        elif(exact_search_num_2 == 1):
            for result in Abbreviation_results:
                # Edge cases
                if ("龙口中集" in i):
                    print('loingkou')
                    result_list.append("Longkou CIMC Raffles Offshore Ltd")
                else:
                    result_list.append(result["Name"])
        else:
            # Perform a Fuzzy search on the database
            Abbreviation_name_2 = "Abbreviation:"+ i
            company_name_2 = "Name:"+ i
            results = solr.search(company_name_2)
            count = 0
            if (len(results) == 0):
                result_list.append("无此公司")
            for result in results:
                if count < 1:
                    # Edge cases
                    if ("龙口" in i):
                        result_list.append("Longkou CIMC Raffles Offshore Ltd")
                    elif("烟台海工" in i):
                        result_list.append("中集海洋工程研究院有限公司")
                    elif("4s"in i):
                        result_list.append("无此公司")
                    else:
                        result_list.append(result["Name"])
                count += 1
    print("result of the list")
    print(len(result_list))
    print("number of weird company")
    print(no_company)

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