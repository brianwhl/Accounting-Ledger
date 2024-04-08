#Tkinter menubar example and connection with Class
from tkinter import*
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Treeview, Combobox
import pymysql
import os, io, sys
from tkinter import Label, Frame, Text, Entry, Canvas,  Button,  StringVar,  filedialog, Menu, messagebox, StringVar, Listbox, Scrollbar
from tkinter import WORD,GROOVE, END, RIDGE,BOTH, VERTICAL, HORIZONTAL,LEFT, RIGHT,NW, X, Y, TOP,BOTTOM, FLAT
import fitz  # PyMuPDF
from PIL.Image import Image
from PIL import Image, ImageTk
from datetime import date, datetime
import datetime
import shutil
import tkinter.font as tkFont
from openpyxl import Workbook
#import subprocess
import pandas as pd
import numpy as np
import customtkinter
from customtkinter import CTkButton, CTkEntry, CTkLabel, CTkFrame, CTkComboBox


class COA(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='powder blue')
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Chart of Account Management System", bd=3, relief=RIDGE, bg="powder blue", 
                        fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)
        
        code_var = StringVar()
        coa1_var = StringVar()
        coa2_var = StringVar()
        curr_var = StringVar()
        cacc_var = StringVar()
        coagroup_var = StringVar()
        memo_var = StringVar()

        #=====Entries Frame=====
        coa_entry_frame=LabelFrame(parent, text="COA Edit", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=10)
        coa_entry_frame.place(x=0, y=55, width=400, height=300)

        lblcode = Label(coa_entry_frame, text="Code", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcoa1 = Label(coa_entry_frame, text="CoaEng", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcoa2 = Label(coa_entry_frame, text="CoaKor", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcurr = Label(coa_entry_frame, text="Currency", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcacc = Label(coa_entry_frame, text="Caccnt", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcoagroup = Label(coa_entry_frame, text="CoaGroup", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblmemo = Label(coa_entry_frame, text="Memo", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")

        lblcode.grid(row=0, column=0, sticky="w")
        lblcoa1.grid(row=1, column=0, sticky="w")
        lblcoa2.grid(row=2, column=0, sticky="w")
        lblcurr.grid(row=3, column=0, sticky="w")
        lblcacc.grid(row=4, column=0, sticky="w")
        lblcoagroup.grid(row=5, column=0, sticky="w")
        lblmemo.grid(row=6, column=0, sticky="w")

        txtcode = Entry(coa_entry_frame, textvariable=code_var, font=("Arial", 12, "bold"), width=30)
        txtcoa1 = Entry(coa_entry_frame, textvariable=coa1_var, font=("Arial", 12, "bold"), width=30)
        txtcoa2 = Entry(coa_entry_frame, textvariable=coa2_var, font=("Arial", 12, "bold"), width=30)
        txtcurr = Entry(coa_entry_frame, textvariable=curr_var, font=("Arial", 12, "bold"), width=30)
        txtcacc = Entry(coa_entry_frame, textvariable=cacc_var, font=("Arial", 12, "bold"), width=30)
        combocoagroup = ttk.Combobox(coa_entry_frame, textvariable=coagroup_var, font=("Arial", 12, "bold"), width=28, state="readonly")
        combocoagroup['values'] = ("Expense", "Sale_Cost","Cash","Asset", "Liablity", "Non_Operation", "ETC")
        combocoagroup.current(0)
        txtmemo = Entry(coa_entry_frame, textvariable=memo_var, font=("Arial", 12, "bold"), width=30)

        txtcode.grid(row=0, column=1)
        txtcoa1.grid(row=1, column=1)
        txtcoa2.grid(row=2, column=1)
        txtcurr.grid(row=3, column=1)
        txtcacc.grid(row=4, column=1)
        combocoagroup.grid(row=5, column=1)
        txtmemo.grid(row=6, column=1)

        # Fetch All Data from DB
        def fetch():
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            cursor.execute("SELECT * FROM coa")
            rows = cursor.fetchall()
            rows=list(rows)
            rows.sort(key=lambda x: x[0], reverse=False)
            mydb.close()
            return rows

        def getData(event):
            selected_row = tv.focus()
            data = tv.item(selected_row)
            global row
            row = data["values"]
            #print(row)
            if len(row) >= 7: #assumming you have 19 columns
                code_var.set(row[0])
                coa1_var.set(row[1])
                coa2_var.set(row[2])
                curr_var.set(row[3])
                cacc_var.set(row[4])
                coagroup_var.set(row[5])
                memo_var.set(row[6])
                # txtAddress.delete(1.0, END)
                # txtAddress.insert(END, row[7])
                
            else:
                messagebox.showerror("Error", "Selected row does not contain enough data.")
            
        def adjustColumnWidths():
            for col in tv["columns"]:
                tv.column(col, width=70)  # Set a default width
                
                # Get the maximum width of data in each column
                max_width = max([len(str(tv.item(item, "values")[int(col) - 1])) for item in tv.get_children()])
                
                # Adjust column width based on maximum data width
                tv.column(col, width=max(20, max_width * 10))  # Minimum width of 70 pixels    

        def dispalyAll():
            tv.delete(*tv.get_children())
            for row in fetch():
                tv.insert("", END, values=row)
            adjustColumnWidths()

        def add_coa():            
            code_var = txtcode.get()
            coa1_var = txtcoa1.get()
            coa2_var = txtcoa2.get()
            curr_var = txtcurr.get()
            cacc_var = txtcacc.get()
            coagroup_var = combocoagroup.get()
            memo_var = txtmemo.get()
            
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            
            if not (coa1_var and coa2_var and coagroup_var):
                messagebox.showerror('Error', 'Enter all fields.')
            else:
                try:
                    sql = "INSERT INTO  coa (CODE, COA1, COA2, CURR, CACC, COAGROUP, MEMO)" \
                        "VALUES (%s,%s, %s, %s, %s, %s, %s)"
                    val = (code_var, coa1_var, coa2_var, curr_var, cacc_var, coagroup_var, memo_var)
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    messagebox.showinfo('Success', 'Data has been inserted.')
                    clearAll()
                    dispalyAll()
                
                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not inserted successfully...") 

        def update_coa():
            selected_item = tv.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to update.')
            else:
                code_var = txtcode.get()
                coa1_var = txtcoa1.get()
                coa2_var = txtcoa2.get()
                curr_var = txtcurr.get()
                cacc_var = txtcacc.get()
                coagroup_var = combocoagroup.get()
                memo_var = txtmemo.get()
                
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "UPDATE coa set COA1=%s, COA2=%s, CURR=%s, CACC=%s, COAGROUP=%s, MEMO=%s where CODE=%s"
                    val = (coa1_var, coa2_var, curr_var, cacc_var, coagroup_var, memo_var, code_var)
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    #add_to_treeview()
                    messagebox.showinfo('Success', 'Data has been updated.')
                
                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not updated successfully...") 
                clearAll()
                dispalyAll()  

        def delete_coa():
            selected_item = tv.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to delete.')
            else:
                code_var = txtcode.get()
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()

                try:
                    sql = "DELETE from coa where CODE=%s"
                    val = (code_var,)
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    # add_to_treeview() 
                    messagebox.showinfo('Success', 'Data has been deleted.')

                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not deleted...")
                clearAll()
                dispalyAll()

        def clearAll():    
            code_var.set("")
            coa1_var.set("")
            coa2_var.set("")
            curr_var.set("")
            cacc_var.set("")
            combocoagroup.set("")
            memo_var.set("")
        
        # =======Button Frame ==============
        btn_frame = Frame(parent, bg="powder blue", bd=3)
        btn_frame.place(x=0, y=358, width=400, height=50)

        btnAdd = Button(btn_frame, command=add_coa, text="Add", width=10, font=("Calibri", 11, "bold"), fg="white",
                        bg="#16a085", bd=0).grid(row=0, column=0, padx=5)
        btnEdit = Button(btn_frame, command=update_coa, text="Update", width=10, font=("Calibri", 11, "bold"),fg="white",
                        bg="#2980b9", bd=0).grid(row=0, column=1, padx=5)
        btnDelete = Button(btn_frame, command=delete_coa, text="Delete", width=10, font=("Calibri", 11, "bold"),fg="white",
                        bg="#c0392b", bd=0).grid(row=0, column=2, padx=5)
        btnClear = Button(btn_frame, command=clearAll, text="Clear", width=10, font=("Calibri", 11, "bold"), fg="white",
                        bg="#f39c12", bd=0).grid(row=0, column=3, padx=5)

        #====Treeview Widget====
        tree_frame=LabelFrame(parent, text="COA DB LIST", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=10)
        tree_frame.place(x=400, y=55, width=1200, height=900)

        style = ttk.Style()
        style.configure("mystyle.Treeview", font=("Arial", 10),rowheight=30)  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=("Arial", 12, "bold"))  # Modify the font of the headings
        tv = ttk.Treeview(tree_frame, columns=(1, 2, 3, 4, 5, 6, 7), style="mystyle.Treeview")
        tv.heading("1", text="Code")
        tv.heading("2", text="Coa1")
        tv.heading("3", text="Coa2")
        tv.heading("4", text="Curr")
        tv.heading("5", text="Cacc")
        tv.heading("6", text="coaGroup")
        tv.heading("7", text="Memo")

        tv['show'] = 'headings'

        # Attach Scrollbars
        xscrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=tv.xview)
        yscrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tv.yview)
        tv.configure(xscrollcommand=xscrollbar.set, yscrollcommand=yscrollbar.set)

        # Grid the Treeview and Scrollbars
        tv.grid(row=0, column=0, sticky=(NSEW))
        xscrollbar.grid(row=1, column=0, sticky=(EW))
        yscrollbar.grid(row=0, column=1, sticky=(NS))

        tv.column("#0", width=0, stretch=YES)

        # Configure grid weights to make Treeview and Scrollbars resize properly
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        tv.bind("<ButtonRelease-1>", getData)
        #tv.pack(fill=BOTH, expand=1)

        dispalyAll()
        
    #====SQLEdit ========================
        def update_sql():
            coagroup_sql = txtcoagroup_1.get()
            code_frm = txtfrm_code.get()
            code_to = txtto_code.get()
        
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
        
            sql = "UPDATE coa SET COAGROUP = %s WHERE CODE BETWEEN %s AND %s"
            val = (coagroup_sql, code_frm, code_to)
            cursor.execute(sql, val)
            mydb.commit()
            mydb.close()
            print("COA updated successfully.")
            dispalyAll()
        
        sql_frame=LabelFrame(parent, text="COA DB EDIT By SQL", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=10)
        sql_frame.place(x=0, y=400, width=400, height=400)
        
        lblfrm_code = Label(sql_frame, text="From", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblto_code = Label(sql_frame, text="To", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcoagroup_1 = Label(sql_frame, text="Group", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblbtn_edit = Label(sql_frame, text="EDIT", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        
        lblfrm_code.grid(row=0, column=0, sticky="w")
        lblto_code.grid(row=1, column=0, sticky="w")
        lblcoagroup_1.grid(row=2, column=0, sticky="w")
        lblbtn_edit.grid(row=3, column=0, sticky="w")

        txtfrm_code = Entry(sql_frame, font=("Arial", 11, "bold"), width=30)
        txtto_code = Entry(sql_frame, font=("Arial", 11, "bold"), width=30)
        txtcoagroup_1 = Entry(sql_frame, font=("Arial", 11, "bold"), width=30)
        btn_edit = Button(sql_frame, text="EDIT", command=update_sql, fg="white", bg="#2980b9", font=("Arial", 11, "bold"), width=26)

        txtfrm_code.grid(row=0, column=1)
        txtto_code.grid(row=1, column=1)
        txtcoagroup_1.grid(row=2, column=1)
        btn_edit.grid(row=3, column=1, sticky="w")
        
#========== FKL ACC ===================#


#========== 여기부터 코드 작성 ==========#

class TBGL(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="powder blue")
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Trial Balance, Ledger Download to Excel file, LineSQ Adjust System", bd=3, relief=RIDGE, 
                        bg="powder blue", fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)
        
        # Variables
        self.start_date = StringVar()
        self.end_date = StringVar()
        self.csv_path = StringVar()
        #self.treeview = ttk.Treeview(treeview_frame)
        
        #-----------------------------------------------
        # Upper Frame
        upperframe = Frame(parent, bd=3, relief=RIDGE, padx=20, bg="powder blue")
        upperframe.place(x=0, y=50, width=1650, height=202)
        
        # Trial Balance Frame/GLedger Download Frame
        tbglframe=LabelFrame(upperframe, text="Date Range", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=2)
        tbglframe.place(x=0, y=2, width=400, height=120)
        
        label_start_date = Label(tbglframe, text="Start Date (YYYY-MM-DD):", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        label_start_date.grid(row=0, column=0, sticky=W)
        entry_start_date = Entry(tbglframe, font=("Arial", 12, "bold"), textvariable=self.start_date, width=20)
        entry_start_date.grid(row=0, column=1, padx=10, pady=10)
        
        label_end_date = Label(tbglframe, text="End Date (YYYY-MM-DD):", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        label_end_date.grid(row=1, column=0, sticky=W)
        entry_end_date = Entry(tbglframe, font=("Arial", 12, "bold"), textvariable=self.end_date, width=20)
        entry_end_date.grid(row=1, column=1, padx=10, pady=10)
        
        # Button Frame
        buttonframe = LabelFrame(upperframe, text="Trial Balance Frame/GLedger Download Frame", bg="powder blue", bd=3, relief=RIDGE, font=("Arial", 12, "bold"), padx=20)
        buttonframe.place(x=405, y=2, width=400, height=120)
        
        btn_trial_balance = Button(buttonframe, text="Trial Balance", command=self.export_trial_balance, font=("Arial", 12, "bold"), width=15, bg="blue", fg="white")
        btn_trial_balance.pack(side=LEFT, padx=10)
        
        btn_ledger = Button(buttonframe, text="Ledger", command=self.export_ledger, font=("Arial", 12, "bold"), width=15, bg="blue", fg="white")
        btn_ledger.pack(side=LEFT, padx=10)
        #-------------------------------------------------
        #Lower Frame
        lowerframe = Frame(parent, bd=3, relief=RIDGE, padx=20, bg="powder blue")
        lowerframe.place(x=0, y=250, width=1650, height=505)
        
        #Create the CSV file upload section
        csv_frame = LabelFrame(lowerframe, text="CSV File Upload")
        csv_frame.place(x=0, y=2, width=400, height=120)

        csv_path = StringVar()
        csv_label = Label(csv_frame, text="PATH:")
        csv_entry = Entry(csv_frame, textvariable=csv_path, width=50)
        csv_button = Button(csv_frame, text="Select Directory", command=self.select_directory)
        upload_button = Button(csv_frame, text="Upload", command=self.import_data)

        csv_label.grid(row=0, column=0, padx=5, pady=5)
        csv_entry.grid(row=0, column=1, padx=5, pady=5)
        csv_button.grid(row=1, column=1, padx=5, pady=5)
        upload_button.grid(row=2, column=1, padx=5, pady=5)

        #Create the update_linesq section
        update_frame = LabelFrame(lowerframe, text="Update Linesq")
        update_frame.place(x=0, y=122, width=400, height=120)

        start_label = Label(update_frame, text="tsq 범위:")
        end_label = Label(update_frame, text="-")

        start_label.grid(row=0, column=0, padx=5, pady=5)
        end_label.grid(row=0, column=2, padx=5, pady=5)

        self.start_entry = Entry(update_frame, width=10)
        self.end_entry = Entry(update_frame, width=10)

        self.start_entry.grid(row=0, column=1, padx=5, pady=5)
        self.end_entry.grid(row=0, column=3, padx=5, pady=5)

        update_button = Button(update_frame, text="Update Linesq", command=self.update_linesq)
        update_button.grid(row=0, column=4, padx=5, pady=5)

        #Create the treeview for displaying data
        treeview_frame = LabelFrame(lowerframe, text="Update Linesq")
        treeview_frame.place(x=403, y=2, width=1200, height=400)
        
        xscroll=Scrollbar(treeview_frame, orient=HORIZONTAL)
        yscroll=Scrollbar(treeview_frame, orient=VERTICAL)

        self.treeview = ttk.Treeview(treeview_frame)
        self.treeview["columns"] = ("tsq","indate", "gubun", "linesq", "voudate", "supplier", "coa", "total","vat", "net", "description", "currency", "vouno", "memo")

        self.treeview.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)

        xscroll.pack(side=BOTTOM, fill=X)
        yscroll.pack(side=RIGHT, fill=Y)
        xscroll.config(command=self.treeview.xview)
        yscroll.config(command=self.treeview.yview)

        self.treeview.column("#0", width=0, stretch=YES)

        for column in self.treeview["columns"]:
            self.treeview.column(column, anchor=CENTER, width=100)
            self.treeview.heading(column, text=column)

            self.treeview.pack(padx=10, pady=10, fill=BOTH, expand=True)

        #==================================================
        
    def export_trial_balance(self):
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        
        try:
            # MySQL 데이터베이스 연결
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            
            # TB
            sql = f"SELECT coa, SUM(net) AS net, SUM(vat) AS vat, SUM(total) AS total FROM gledger WHERE voudate BETWEEN '{start_date}' AND '{end_date}' GROUP BY coa"
            cursor.execute(sql)
            result = cursor.fetchall()
            
            # 결과를 데이터프레임으로 변환
            df = pd.DataFrame(result, columns=['coa', 'net', 'vat', 'total'])
            
            # 엑셀 파일 저장 경로 설정
            today = datetime.today().strftime('%Y%m%d')
            file_path = f"D:\\15Work\\fkldb\\2023\\{today}_TB_{start_date}_{end_date}.xlsx"
            
            # 엑셀 파일로 저장
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Trial Balance data has been saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            mydb.close()
    
    def export_ledger(self):
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        
        try:
            # MySQL 데이터베이스 연결
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            
            # Gledger All
            sql = f"SELECT * FROM gledger WHERE voudate BETWEEN '{start_date}' AND '{end_date}' ORDER BY voudate"
            cursor.execute(sql)
            result = cursor.fetchall()
            
            # 결과를 데이터프레임으로 변환
            df = pd.DataFrame(result, columns=['tsq', 'indate', 'gubun', 'linesq', 'voudate', 'supplier', 'coa', 'total', 'vat', 'net', 'description', 'currency', 'vouno', 'memo'])
            
            # 엑셀 파일 저장 경로 설정
            today = datetime.today().strftime('%Y%m%d')
            file_path = f"D:\\15Work\\fkldb\\2023\\{today}_gl_{start_date}_{end_date}.xlsx"
            
            # 엑셀 파일로 저장
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Ledger data has been saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            mydb.close()
            
    #---------------------------------------------------
    def import_data(self):
        csv_file = self.csv_path.get() 
        if csv_file:
            # Connect to the MySQL database
            conn = pymysql.connect(host='localhost', user='root', password='0000', db='fkl')
            # Read the CSV file into a pandas DataFrame
            df = pd.read_csv(csv_file, skiprows=1, na_values=["nan", "NaN", ""])
            # Option 1: Replace NaN with NULL
            df.replace(np.nan, None, inplace=True)
            # Get column names in the correct order (assuming here)
            column_names = ["indate", "gubun", "linesq", "voudate", "supplier", "coa", "total",
                            "vat", "net", "description", "currency", "vouno", "memo"]
            # Prepare SQL query dynamically using column names
            query = "INSERT INTO gledger ({}) VALUES ({})".format(
                ",".join(column_names), ",".join(["%s"] * len(column_names))
            )
            # Execute query with executemany for efficiency
            cursor = conn.cursor()
            cursor.executemany(query, df.values.tolist())
            conn.commit()
            # Close connections
            cursor.close()
            conn.close()
            print("Data imported successfully!")
            self.display_data()


    def update_linesq(self):
        start_tsq = int(self.start_entry.get())
        end_tsq = int(self.end_entry.get())

        mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = mydb.cursor()
        for tsq in range(start_tsq, end_tsq + 1):
            # linesq 값 가져오기
            query = "SELECT linesq FROM gledger WHERE tsq = %s"
            cursor.execute(query, (tsq,))
            result = cursor.fetchone()
            if result:
                linesq = result[0]
                if linesq:
                    # '-' 앞의 값 추출
                    linesq_value = linesq.split('-')[0]
        
                    # tsq와 linesq 값 비교
                    if int(linesq_value) != tsq:
                        # linesq 값 업데이트
                        new_linesq = str(tsq) + '-' + linesq.split('-')[1]
                        update_query = "UPDATE gledger SET linesq = %s WHERE tsq = %s"
                        cursor.execute(update_query, (new_linesq, tsq))
                        mydb.commit()
                        print(f"tsq: {tsq}, linesq: {linesq} -> {new_linesq} (updated)")
                    else:
                        print(f"tsq: {tsq}, linesq: {linesq} (no update needed)")
                else:
                    print(f"tsq: {tsq}, linesq: None (no update needed)")
            else:
                print(f"tsq: {tsq} not found in the table")
            
        # treeview의 커서를 입력된 범위의 tsq row로 이동
        #self.treeview = ttk.Treeview(self) # Add missing initialization
        for child in self.treeview.get_children():
            values = self.treeview.item(child)["values"]
            try:
                tsq = int(values[self.treeview["columns"].index("linesq")].split("-")[0])  # linesq 열에서 tsq 값 추출
                if start_tsq <= tsq <= end_tsq:
                    self.treeview.selection_set(child)
                    self.treeview.focus(child)
                    self.treeview.see(child)
                    break
            except (IndexError, ValueError):
                continue

        cursor.close()
        mydb.close()
        self.display_data()


    def select_directory(self):
        directory = filedialog.askdirectory()
        self.csv_path.set(directory)

    def display_data(self):
        # Clear existing data in the treeview
        for row in self.treeview.get_children():
            self.treeview.delete(row)
        
        # Fetch data from the database
        conn = pymysql.connect(host='localhost', user='root', password='0000', db='fkl')
        cursor = conn.cursor()
        query = "SELECT * FROM gledger"
        cursor.execute(query)
        data = cursor.fetchall()

        # Insert data into the treeview
        for row in data:
            self.treeview.insert("", "end", values=row)

        cursor.close()
        conn.close()
        
        # Scroll to the row with the specified value
        start_tsq = int(self.start_entry.get())
        self.scroll_to_value(start_tsq)

    def scroll_to_value(self, value):
        # Find the row with the specified value
        for row in self.treeview.get_children():
            if self.treeview.item(row)['values'][0] == value:
                # Scroll to the found row
                self.treeview.see(row)
                break


#========== 프레임 코드 작성 ==========#
class FKL(tk.Frame):  
    def __init__(self, parent):
        super().__init__(parent, bg='powder blue')
        self.pack(fill=tk.BOTH, expand=True)
        label = tk.Label(self, text="FKL System")
        label.pack()

#========== 메뉴바 코드 작성 ==========#
class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("메인창")
        self.geometry("2570x1100+0+0")

        menubar = Menu(self)
        self.config(menu=menubar,bg='powder blue') #'cadet blue'

        #-----------------기본메뉴-----------------#
        basicmenu=Menu(menubar, tearoff=0) 
        menubar.add_cascade(label='BASIC', menu=basicmenu)
        basicmenu.add_command(label='COA', command=self.open_coa, font=('Arial', 11, 'bold'))
        basicmenu.add_command(label='Customer', command=self.open_customer, font=('Arial', 11, 'bold'))
        basicmenu.add_command(label='Employee', command=self.open_employee, font=('Arial', 11, 'bold'))
        basicmenu.add_separator()
        basicmenu.add_command(label='종료', command=self.destroy)
        
        #-----------------데이터메뉴-----------------#
        datamenu=Menu(menubar, tearoff=0)
        menubar.add_cascade(label='DATA', menu=datamenu) #데이터메뉴
        datamenu.add_command(label='ScanDocs', command="", font=('Arial', 11, 'bold'))
        datamenu.add_command(label='PNG->Excel', command="", font=('Arial', 11, 'bold'))
        datamenu.add_command(label='Excel->CSV', command="", font=('Arial', 11, 'bold'))
        datamenu.add_command(label='Upload To DB', command="", font=('Arial', 11, 'bold'))
        datamenu.add_command(label='Image Save To DB', command=self.open_imagesavedb, font=('Arial', 11, 'bold'))
        datamenu.add_separator()
        datamenu.add_command(label='TB/GL DOWNLOAD, AJE', command=self.open_tbgl)

        #-----------------계정작업-----------------#
        glmenu=Menu(menubar, tearoff=0)
        menubar.add_cascade(label='LEDGER', menu=glmenu) #일반메뉴
        glmenu.add_command(label='Sales Invoice', command=self.open_invoice, font=('Arial', 11, 'bold'))
        glmenu.add_command(label='Purchasing', command=self.open_purchase, font=('Arial', 11, 'bold'))
        glmenu.add_command(label='Expenses', command=self.open_receipt, font=('Arial', 11, 'bold'))
        glmenu.add_separator()
        transmenu=Menu(glmenu, tearoff=0)
        glmenu.add_cascade(label='TRANSACTION', menu=transmenu)
        transmenu.add_command(label='Other', command=self.open_Aledger, font=('Arial', 11, 'bold'))
        transmenu.add_command(label='Depriciation', command="", font=('Arial', 11, 'bold'))
        transmenu.add_command(label='General Journal', command="", font=('Arial', 11, 'bold'))
        
        #-----------------원장기록-----------------#
        postmenu=Menu(menubar, tearoff=0)
        menubar.add_cascade(label='POSTING', menu=postmenu)
        
        postmenu.add_command(label='YearEnd Input', command="", font=('Arial', 11, 'bold'))
        postmenu.add_command(label='Jounalizing', command="", font=('Arial', 11, 'bold'))
        postmenu.add_command(label='Ledger Export', command="", font=('Arial', 11, 'bold'))
        postmenu.add_command(label='Financial Statement', command="", font=('Arial', 11, 'bold'))
        
        #-----------------기초화면-----------------#
        menubar.add_command(label='FKL', command=self.open_fkl)
        
        menubar.add_command(label='종료', command=self.destroy)

        self.window1 = None
        self.window2 = None
        self.window3 = None  
        self.window4 = None
        self.window5 = None
        self.window6 = None
        self.window7 = None
        self.window8 = None
        self.window9 = None
        self.window10 = None

    def open_coa(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window1 = COA(self)
        
    def open_receipt(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window2 = Receipt(self)

    def open_customer(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window3 = Customer(self)
        
    def open_employee(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window4 = Employee(self)

    def open_imagesavedb(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window5 = ImageSaveDB(self)
        
    def open_Aledger(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window6 = Aledger(self)

    def open_invoice(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window7 = Invoice(self)
    
    def open_purchase(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window8 = Purchase(self)
    
    def open_tbgl(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window9 = TBGL(self)
        
    def open_fkl(self):
        self.close_all_windows()  # 모든 창을 닫음
        self.window10 = FKL(self)
        
        
        

    def close_all_windows(self):
        if self.window1 is not None:
            self.window1.destroy()
            self.window1 = None
        if self.window2 is not None:
            self.window2.destroy()
            self.window2 = None
        if self.window3 is not None:
            self.window3.destroy()
            self.window3 = None
        if self.window4 is not None:
            self.window4.destroy()
            self.window4 = None
        if self.window5 is not None:
            self.window5.destroy()
            self.window5 = None
        if self.window6 is not None:
            self.window6.destroy()
            self.window6 = None
        if self.window7 is not None:
            self.window7.destroy()
            self.window7 = None
        if self.window8 is not None:
            self.window8.destroy()
            self.window8 = None
        if self.window9 is not None:
            self.window9.destroy()
            self.window9 = None
        if self.window10 is not None:
            self.window10.destroy()
            self.window10 = None


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
