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
class Receipt(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, background="cadet blue")
        self.label = Label(self, text="Flowconn Korea Inc Accounting system")
        self.label.grid(row=0, column=0, padx=5, pady=10)
        self.pack(fill=tk.BOTH, expand=True)  # Expand the frame to fill the parent
        #self.config(bg="cadet blue")
        
        today = date.today().strftime("%Y-%m-%d")
        
        #tsq자동증가 번호메기기
        con=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
        cur=con.cursor()
        cur1=con.cursor()
        
        # Get the maximum value of the tsq field from the gledger table
        cur.execute("SELECT MAX(tsq) FROM gledger")
        result = cur.fetchone()
        if result is not None:
            max_tsq = result[0]
        else:
            max_tsq = 0  # Assign a default value of 0 if max_tsq is None

        self.tsq_var = StringVar()
        self.tsq_var.set(str(max_tsq + 1))
        self.indate_var = StringVar()
        self.indate_var.set(today)
        self.gubun_var = StringVar()
        self.linesq_var = StringVar()
        self.linesq_var.set(str(self.tsq_var.get()) + "-1")
        self.voudate_var = StringVar()
        self.voudate_var.set(today)
        self.supplier_var = StringVar()
        self.coa_var = StringVar()
        self.total_var = StringVar()
        self.total_var.trace('w', self.calculate_net)
        self.vat_var = StringVar()
        self.vat_var.trace('w', self.calculate_net)
        self.net_var = StringVar()

        # 이전에 'float' 값을 설정한 코드를 수정하여 'StringVar'로 변환합니다.
        self.total_var.set(str(0.0))
        self.vat_var.set(str(0.0))

        # # calculate_net 메소드를 호출하여 var_net 값을 계산합니다.
        # self.calculate_net()
        # # trace() 메소드를 사용하여 var_total과 var_vat 값이 변경될 때마다 calculate_net() 메소드를 호출하도록 설정합니다.
        # self.total_var.trace('w', self.calculate_net)
        # self.vat_var.trace('w', self.calculate_net)
        
        self.description_var = StringVar()
        self.currency_var = StringVar()
        self.vouno_var = StringVar()
        
        self.search_by=StringVar()
        self.search_txt = StringVar()
        
        #global scale
        self.scale = 1.0
        self.photo = None  # Initialize photo as None
        self.xx_var=1.0

        #==============Combobox_Frame============================================
        combobox_Frame = Frame(self,bd=4,relief=FLAT,bg="antiquewhite2")
        combobox_Frame.place(x=5,y=60,width=600,height=150) #(x=20,y=100,width=450,height=1000)+1250

        combobox_title = Label(combobox_Frame,text="Scanned Receipt:", fg="black",font=("Arial",12,"bold"))
        combobox_title.grid(row=0, column=0, pady=0, sticky="w")

        # Get a list of all files in the specified directory
        receipt_folder = "D:\\15Work\\Scanner\\2023"
        receipt_files = [f for f in os.listdir(receipt_folder)]

        # Populate the combobox with file names
        self.pdf_combobox = Combobox(combobox_Frame, values=receipt_files, width=30,font=("Arial",12,"bold"))
        self.pdf_combobox['values'] = receipt_files
        self.pdf_combobox.grid(row=0, column=0, padx=150, sticky="w")
        #self.pdf_combobox.bind("<<ComboboxSelected>>", self.display_file)

        # AAA Additional Textbox to show path and file name
        self.file_info_textbox = Text(combobox_Frame, wrap=WORD, height=2, width=45,font=("Arial",12,"bold"))
        self.file_info_textbox.grid(row=1,column=0,padx=150, sticky="w")

        # Update txt_vouno based on the selected file
        #self.update_vouno_from_filename()
        fileNameChange_btn = Button(combobox_Frame, text="FileNameChange", width=15,command=self.update_vouno_from_filename)
        fileNameChange_btn.grid(row=1, column=0, padx=770, sticky="w")
        
        # Update the list of values when a PDF is selected
        #self.pdf_combobox.bind("<<ComboboxSelected>>", self.update_file_info_textbox)
        self.pdf_combobox.bind("<<ComboboxSelected>>", self.display_file)
        
        #==============Display_Frame==================================================
        # left2_frame - File Display
        file_display_frame = Frame(self,bd=4,relief=FLAT,bg="#FBD96C")
        file_display_frame.place(x=5,y=220,width=900,height=838)

        # Call display_file with the initial PDF file path
        initial_pdf_path = os.path.join(receipt_folder, self.pdf_combobox.get())
        self.display_file(initial_pdf_path)

        self.canvas = Canvas(file_display_frame)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scale = 1
        
        # Set up event bindings for zooming
        #self.canvas.bind("<B1-Motion>", self.zoom)
        #self.canvas.bind("<MouseWheel>", self.zoom)

        # Scrollbar for vertical scrolling
        self.scrollbar = ttk.Scrollbar(file_display_frame, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Bind mousewheel scrolling
        self.canvas.bind_all("<<MouseWheel>>", self.on_mousewheel)

        # Buttons to set different scales
        btn1 = Button(combobox_Frame, text="Btn2", command=lambda: self.set_value(2.0))
        btn2 = Button(combobox_Frame,text="Btn1.5", command=lambda: self.set_value(1.5))
        btn3 = Button(combobox_Frame,text="Btn3/4", command=lambda: self.set_value(0.75))
        btn4 = Button(combobox_Frame,text="Btn1/2", command=lambda: self.set_value(0.5))
        btn5 = Button(combobox_Frame,text="Btn1", command=lambda: self.set_value(1.0))

        btn1.grid(row=2,column=0,padx=50, pady=5, sticky="w")
        btn2.grid(row=2,column=0,padx=100, pady=5, sticky="w")
        btn3.grid(row=2,column=0,padx=150, pady=5, sticky="w")
        btn4.grid(row=2,column=0,padx=200, pady=5, sticky="w")
        btn5.grid(row=2,column=0,padx=250, pady=5, sticky="w")

        #==============Receipt_input_Frame============================================
        r_input_frame = Frame(self,bd=4,relief=RIDGE,bg='honeydew3')
        r_input_frame.place(x=930,y=60,width=450,height=1000) #(x=20,y=100,width=450,height=1000)+1250

        m_title = Label(r_input_frame,text="Receipt Input",bg='honeydew3',fg="black",font=("Malgun Gothic",12,"bold"))
        m_title.grid(row=0 ,columnspan=2,pady=20)

        Resetbtn = Button(r_input_frame, text="Reset", width=10,command=self.reset).grid(row=0, column=1, padx=20, pady=20, sticky="e")

        lbl_tsq = Label(r_input_frame,text="TSQ:", bg='honeydew3',fg="white",font=("Malgun Gothic",13,"bold"))
        lbl_tsq.grid(row=1 ,column=0,pady=10,padx=20,sticky="w")
        txt_tsq= Entry(r_input_frame,textvariable=self.tsq_var,font=("Malgun Gothic",13,"bold"), width=24, bd=5, relief=FLAT)
        txt_tsq.grid(row=1 ,column=1,pady=10,padx=20,sticky="w")

        lbl_indate = Label(r_input_frame, text="INPUT DATE:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_indate.grid(row=2, column=0, pady=10, padx=20, sticky="w")
        txt_indate = Entry(r_input_frame,textvariable=self.indate_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_indate.grid(row=2, column=1, pady=10, padx=20, sticky="w")

        lbl_gubun = Label(r_input_frame, text="GUBUN:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_gubun.grid(row=3, column=0, pady=10, padx=20, sticky="w")
        combo_gubun = ttk.Combobox(r_input_frame,textvariable=self.gubun_var,font=("Malgun Gothic", 13, "bold"),width=23, state='normal')
        combo_gubun['values'] = ("Expenses","Bank","Purchase","Sales","Asset","Lability", "Capital")
        combo_gubun.current(0)
        combo_gubun.grid(row=3,column=1,pady=10,padx=20,sticky="w")

        lbl_linesq = Label(r_input_frame, text="LINE SQ_NO:",bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_linesq.grid(row=4, column=0, pady=10, padx=20, sticky="w")
        txt_linesq = Entry(r_input_frame, textvariable=self.linesq_var,font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_linesq.grid(row=4, column=1, pady=10, padx=20, sticky="w")

        lbl_voudate = Label(r_input_frame, text="RECEIPT DATE:",bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_voudate.grid(row=5, column=0, pady=10, padx=20, sticky="w")
        txt_voudate = Entry(r_input_frame, textvariable=self.voudate_var,font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_voudate.grid(row=5, column=1, pady=10, padx=20, sticky="w")

        lbl_supplier = Label(r_input_frame, text="SUPPLIER:",bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_supplier.grid(row=6, column=0, pady=10, padx=20, sticky="w")
        txt_supplier = Entry(r_input_frame, textvariable=self.supplier_var,font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_supplier.grid(row=6, column=1, pady=10, padx=20, sticky="w")

        lbl_coa = Label(r_input_frame, text="COA SELECT:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_coa.grid(row=7, column=0, pady=10, padx=20, sticky="w")
        cur.execute("select COA2 from coa")
        coa_data = cur.fetchall()
        coa_values = [row[0] for row in coa_data]
        combo_coa = ttk.Combobox(r_input_frame,textvariable=self.coa_var,font=("Malgun Gothic", 13, "bold"), width=23 ,state='readonly')
        combo_coa['values'] = coa_values
        combo_coa.current(53)
        combo_coa.grid(row=7,column=1,pady=10,padx=20,sticky="w")

        lbl_total = Label(r_input_frame, text="TOTAL:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_total.grid(row=8, column=0, pady=10, padx=20, sticky="w")
        txt_total = Entry(r_input_frame,textvariable=self.total_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_total.grid(row=8, column=1, pady=10, padx=20, sticky="w")
        
        lbl_vat = Label(r_input_frame, text="VAT:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_vat.grid(row=9, column=0, pady=10, padx=20, sticky="w")
        txt_vat = Entry(r_input_frame,textvariable=self.vat_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_vat.grid(row=9, column=1, pady=10, padx=20, sticky="w")
        
        lbl_net = Label(r_input_frame, text="NET:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_net.grid(row=10, column=0, pady=10, padx=20, sticky="w")
        txt_net = Entry(r_input_frame,textvariable=self.net_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_net.grid(row=10, column=1, pady=10, padx=20, sticky="w")
        
        lbl_description = Label(r_input_frame, text="DESCRIPTION:",bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_description.grid(row=11, column=0, pady=10, padx=20, sticky="w")
        txt_description = Entry(r_input_frame,textvariable=self.description_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_description.grid(row=11, column=1, pady=10, padx=20, sticky="w")

        lbl_currency = Label(r_input_frame, text="CURRENCY:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_currency.grid(row=12, column=0, pady=10, padx=20, sticky="w")
        txt_currency = Entry(r_input_frame,textvariable=self.currency_var, font=("Malgun Gothic", 13, "bold"), width=24, bd=5, relief=FLAT)
        txt_currency.grid(row=12, column=1, pady=10, padx=20, sticky="w")

        lbl_vouno = Label(r_input_frame, text="VOU_NO:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_vouno.grid(row=13, column=0, pady=10, padx=20, sticky="w")
        txt_vouno = Entry(r_input_frame,textvariable=self.vouno_var, font=("Malgun Gothic", 10, "bold"), width=24, bd=5, relief=FLAT)
        txt_vouno.grid(row=13, column=1, pady=10, padx=20, sticky="w")

        lbl_memo = Label(r_input_frame, text="MEMO:", bg='honeydew3', fg="white", font=("Malgun Gothic", 13, "bold"))
        lbl_memo.grid(row=14, column=0, pady=10, padx=20, sticky="w")
        self.txt_memo = Text(r_input_frame, width=25,height=3,font=("Malgun Gothic",13,"bold"))
        self.txt_memo.grid(row=14, column=1, pady=10, padx=20, sticky="w")

#=========Button Frame==================
        btn_Frame = Frame(r_input_frame, bd=3, relief=RIDGE, bg='honeydew3')
        btn_Frame.place(x=12, y=870, width=420) #(x=15, y=925, width=420)+1250

        Addbtn = Button(btn_Frame,text="Add",width=10,command=self.add_supplier).grid(row=0,column=0,padx=10,pady=10)
        updatebtn = Button(btn_Frame, text="Update", width=10,command=self.update_data).grid(row=0, column=1, padx=10, pady=10)
        deletebtn = Button(btn_Frame, text="Delete", width=10,command=self.delete_data).grid(row=0, column=2, padx=10, pady=10)
        Clearbtn = Button(btn_Frame, text="Clear", width=10,command=self.clear).grid(row=0, column=3, padx=10, pady=10)

# =========2nd Detials  Frame==================
        Detials_Frame = Frame(self, bd=4, relief=RIDGE, bg="#6D7434")
        Detials_Frame.place(x=1385, y=60, width=1175, height=1000) #(x=500, y=100, width=800, height=585)+1250

        lbl_search = Label(Detials_Frame, text="Search By", bg="#6D7434", fg="white",font=("Malgun Gothic", 13, "bold"))
        lbl_search.grid(row=0, column=0, pady=10, padx=20, sticky="w")

        combo_search = ttk.Combobox(Detials_Frame,textvariable=self.search_by,width=10, font=("Malgun Gothic", 13, "bold"), state='readonly')
        combo_search['values'] = ("coa", "voudate", "gubun")
        combo_search.grid(row=0, column=1, padx=20, pady=10)

        txt_search= Entry(Detials_Frame,textvariable=self.search_txt,width=20, font=("Malgun Gothic", 10, "bold"), bd=5, relief=GROOVE)
        txt_search.grid(row=0, column=2, pady=10, padx=20, sticky="w")

        searchbtn = Button(Detials_Frame, text="Search", width=10,pady=5,command=self.search_data).grid(row=0, column=3, padx=10, pady=10)
        showallbtn = Button(Detials_Frame, text="Show All", width=10, pady=5,command=self.fetch_data).grid(row=0, column=4, padx=10, pady=10)

#========== table frame ===========
        Table_Frame = Frame(Detials_Frame, bd=4, relief=RIDGE, bg="crimson")
        Table_Frame.place(x=6, y=70, width=1160, height=910)

        scroll_x = ttk.Scrollbar(Table_Frame,orient=HORIZONTAL)
        scroll_y = ttk.Scrollbar(Table_Frame,orient=VERTICAL)
        self.gledger_table = ttk.Treeview(Table_Frame,columns=("tsq","indate","gubun","linesq","voudate","supplier","coa","total","vat","net","description","currency","vouno","memo"),
                                        xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM,fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
        scroll_x.config(command=self.gledger_table.xview)
        scroll_y.config(command=self.gledger_table.yview)
        self.gledger_table.heading("tsq",text="SQNO.")
        self.gledger_table.heading("indate", text="InputDate")
        self.gledger_table.heading("gubun", text="GUBUN")
        self.gledger_table.heading("linesq", text="LINENO")
        self.gledger_table.heading("voudate", text="RDATE")
        self.gledger_table.heading("supplier", text="SUPPLIER")
        self.gledger_table.heading("coa", text="COANO")
        self.gledger_table.heading("total", text="TOTAL")
        self.gledger_table.heading("vat", text="VAT")
        self.gledger_table.heading("net", text="NET")
        self.gledger_table.heading("description", text="DESCRIPTION")
        self.gledger_table.heading("currency", text="CURRENCY")
        self.gledger_table.heading("vouno", text="VOUNO")
        self.gledger_table.heading("memo", text="MEMO")
        
        self.gledger_table['show'] = 'headings'
        self.gledger_table.column("tsq",width=100)
        self.gledger_table.column("indate", width=100)
        self.gledger_table.column("gubun", width=100)
        self.gledger_table.column("linesq", width=100)
        self.gledger_table.column("voudate", width=100)
        self.gledger_table.column("supplier", width=100)
        self.gledger_table.column("coa", width=100)
        self.gledger_table.column("total", width=100)
        self.gledger_table.column("vat", width=100)
        self.gledger_table.column("net", width=100)
        self.gledger_table.column("description", width=100)
        self.gledger_table.column("currency", width=100)
        self.gledger_table.column("vouno", width=100)
        self.gledger_table.column("memo", width=100)

        self.gledger_table.pack(fill=BOTH , expand=1)
        self.gledger_table.bind("<ButtonRelease-1>",self.get_cursor)

        self.fetch_data()
            
    # def update_file_info_textbox(self, event):
    #     selected_file = self.pdf_combobox.get()
    #     file_info = os.path.join("D:\\fkldb\\2023", selected_file)
    #     self.file_info_textbox.delete(1.0, END)
    #     self.file_info_textbox.insert(END, file_info)
    
    def update_vouno_from_filename(self):
        vouno_from_filename = self.file_info_textbox.get(1.0, END).strip()
        self.vouno_var.set(vouno_from_filename)

    def display_file(self, event):
        selected_file = self.pdf_combobox.get()
        file_info = os.path.join("D:\\15Work\\Scanner\\2023", selected_file)
        file_info1 = os.path.join("D:\\15Work\\fkldb\\2023", selected_file)
        
        # Check if file_info_textbox is properly initialized
        if hasattr(self, 'file_info_textbox') and self.file_info_textbox:
            self.file_info_textbox.delete(1.0, END)
            self.file_info_textbox.insert(END, file_info)

        if selected_file.lower().endswith('.pdf'):
            self.show_pdf(file_info)
        elif selected_file.lower().endswith(('.png', '.jpg', '.jpeg')):
            self.show_image(file_info)
            
        # Set file path and file name to txt_vouno
        self.vouno_var.set(file_info1)

    def show_pdf(self, pdf_path):
        pass

    def show_image(self, image_path):
        try:
            image = Image.open(image_path)
            new_size = (int(image.width * self.xx_var), int(image.height * self.xx_var))
            resized_image = image.resize(new_size)
            photo = ImageTk.PhotoImage(resized_image)
            self.canvas.config(scrollregion=(0, 0, new_size[0], new_size[1]), width=new_size[0], height=new_size[1])
            self.canvas.create_image(0, 0, anchor=NW, image=photo)
            #self.canvas.image= photo  # keep reference to image
            self.photo = photo  # keep reference to image
            
        except Exception as e:
            messagebox.showerror("Error", f"Error displaying image: {str(e)}")

    def set_value(self, value):
        self.xx_var=float(value)
        print(f"Button clicked! Value set to {value}")
        selected_file = self.pdf_combobox.get()
        file_info = os.path.join("D:\\15Work\\Scanner\\2023", selected_file)
        self.display_file(file_info)

    def on_mousewheel(self, event):
        # Enable scrolling with the mouse wheel
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")
        self.canvas.xview_scroll(-1 * (event.delta // 120), "units")

    def max_tsq_new(self):
        #tsq자동증가 번호메기기
        con = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cur2 = con.cursor()
        
        # Get the maximum value of the tsq field from the gledger table
        cur2.execute("SELECT MAX(tsq) FROM gledger")
        result = cur2.fetchone()
        max_tsq = result[0] if result is not None else 0
        self.tsq_var.set(str(max_tsq+1))
        self.linesq_var.set(str(max_tsq+1) + "-1")
        
        # Get the maximum value of the vouno field from the gledger table
        # cur2.execute("SELECT MAX(vouno) FROM gledger")
        # max_vouno = cur2.fetchone()[0]
        # self.vouno_var.set(max_vouno+1)
                
        # self.vat_var.set(0)
        # self.currency_var.set("krw")
    
    #self.var_net계산
    def calculate_net(self, *args):
            try:
                total = float(self.total_var.get())
                vat = float(self.vat_var.get())
                net = total-vat
                self.net_var.set(str(net))
                # 순수익 라벨 업데이트
                #self.lbl_net.config(text="{:.2f}".format(due))
            except ValueError:
                pass 

    def add_supplier(self):
        if self.tsq_var.get()=="" or self.indate_var.get()=="" :
            messagebox.showerror("Error","All fields are requried")
        else:
            con=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
            cur=con.cursor()

            cur.execute("insert into gledger values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(self.tsq_var.get(),
                                                                        self.indate_var.get(),
                                                                        self.gubun_var.get(),
                                                                        self.linesq_var.get(),
                                                                        self.voudate_var.get(),
                                                                        self.supplier_var.get(),
                                                                        self.coa_var.get(),
                                                                        self.total_var.get(),
                                                                        self.vat_var.get(),
                                                                        self.net_var.get(),
                                                                        self.description_var.get(),
                                                                        self.currency_var.get(),
                                                                        self.vouno_var.get(),
                                                                        self.txt_memo.get('1.0',END)
                                                                        ))

            con.commit()
            self.fetch_data()
            self.clear()
            con.close()
            messagebox.showinfo("Success","Record has been inserted")
            
            #Moved from Source to Destination
            selected_file = self.pdf_combobox.get()
            source_path = os.path.join("D:\\15Work\\Scanner\\2023", selected_file)
            destination_path = "D:\\15Work\\fkldb\\2023"

            try:
                # Move the selected file to the destination directory
                shutil.move(source_path, destination_path)
                messagebox.showinfo("Success", f"File {selected_file} moved to {destination_path}")
            except FileNotFoundError:
                messagebox.showerror("Error", f"File {selected_file} not found.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

            # Refresh the combobox to reflect the updated file list
            self.refresh_combobox()
    
    def refresh_combobox(self):
        # Update the combobox with the list of PDF files in the source directory
        pdf_files = [file for file in os.listdir("D:\\15Work\\Scanner\\2023")] # if file.lower().endswith('.pdf')]
        self.pdf_combobox['values'] = pdf_files
        self.pdf_combobox.set('')  # Clear the selection

        # Update the file info textbox and display the first PDF if available
        #self.update_file_info_textbox(None)
        self.display_file(None)
                
    def fetch_data(self):
        con = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cur = con.cursor()
        cur.execute("select * from gledger ORDER BY tsq DESC")
        rows = cur.fetchall()
        if len(rows)!=0:
            self.gledger_table.delete(*self.gledger_table.get_children())
            for row in rows:
                self.gledger_table.insert('',END,values=row)
            con.commit()
        con.close()

    def clear(self):
        self.tsq_var.set("")
        self.indate_var.set("")
        self.gubun_var.set("")
        self.linesq_var.set("")
        self.voudate_var.set("")
        self.supplier_var.set("")
        self.coa_var.set("")
        self.total_var.set("")
        self.vat_var.set("")
        self.net_var.set("")
        self.description_var.set("")
        self.currency_var.set("")
        self.vouno_var.set("")
        self.txt_memo.delete("1.0",END)

    def get_cursor(self, ev):
        cursor_row = self.gledger_table.focus()
        contents = self.gledger_table.item(cursor_row)
        row = contents['values']
        self.tsq_var.set(row[0])
        self.indate_var.set(row[1])
        self.gubun_var.set(row[2])
        self.linesq_var.set(row[3])
        self.voudate_var.set(row[4])
        self.supplier_var.set(row[5])
        self.coa_var.set(row[6])
        self.total_var.set(row[7])
        self.vat_var.set(row[8])
        self.net_var.set(row[9])
        self.description_var.set(row[10])
        self.currency_var.set(row[11])
        self.vouno_var.set(row[12])
        self.txt_memo.delete("1.0", END)
        self.txt_memo.insert(END, row[13])

    def update_data(self):
        con = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cur = con.cursor()
        cur.execute("update gledger set indate=%s,gubun=%s,linesq=%s,voudate=%s,supplier=%s,coa=%s,total=%s,vat=%s,net=%s,description=%s,currency=%s,vouno=%s,memo=%s where tsq=%s", (
                                                                        self.indate_var.get(),
                                                                        self.gubun_var.get(),
                                                                        self.linesq_var.get(),
                                                                        self.voudate_var.get(),
                                                                        self.supplier_var.get(),
                                                                        self.coa_var.get(),
                                                                        self.total_var.get(),
                                                                        self.vat_var.get(),
                                                                        self.net_var.get(),
                                                                        self.description_var.get(),
                                                                        self.currency_var.get(),
                                                                        self.vouno_var.get(),
                                                                        self.txt_memo.get('1.0', END),
                                                                        self.tsq_var.get()
                                                                        ))

        con.commit()
        self.fetch_data()
        self.clear()
        con.close()

    def delete_data(self):
        con = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cur = con.cursor()
        cur.execute("delete from gledger where tsq=%s",self.tsq_var.get())
        con.commit()
        con.close()
        self.fetch_data()
        self.clear()

    def reset(self):
        today = date.today().strftime("%Y-%m-%d")
        self.refresh_combobox()
        self.tsq_var.get(),
        self.indate_var.set(today),
        self.gubun_var.get(),
        #self.linesq_var.set(""),
        self.voudate_var.set(today),
        self.supplier_var.set(""),
        self.coa_var.get(),
        #self.total_var.set(""),
        #self.vat_var.set(""),
        #self.net_var.set(""),
        self.description_var.set(""),
        self.currency_var.set("krw"),
        self.vouno_var.get()
        self.max_tsq_new()
        self.fetch_data()

    def search_data(self):
        con = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cur = con.cursor()
        cur.execute("select * from gledger where " + str(self.search_by.get())+" Like '%"+str(self.search_txt.get())+"%'")
        rows = cur.fetchall()
        if len(rows)!=0:
            self.gledger_table.delete(*self.gledger_table.get_children())
            for row in rows:
                self.gledger_table.insert('',END,values=row)
            con.commit()
        con.close()

#========== 여기까지 코드 작성 ==========#
class Customer(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent,background="cadet blue")
        self.label = tk.Label(self, text="Customer system")
        self.label.grid(row=0, column=0, padx=5, pady=10)
        self.pack(fill=tk.BOTH, expand=True)  # Expand the frame to fill the parent
        #self.config(bg="cadet blue")

        wrapper1 = tk.LabelFrame(text="Search", width=2000, height=70)
        wrapper2 = tk.LabelFrame(text="Customer Data", width=2000, height=170)
        wrapper3 = tk.LabelFrame(text="Customer List")

        wrapper1.place(x=6, y=30)
        wrapper2.place(x=6, y=80)
        wrapper3.place(x=6, y=200, width=2000, height=700)


        q = StringVar()
        t1 = StringVar()
        t2 = StringVar()
        t3 = StringVar()
        t4 = StringVar()
        t4 = StringVar()
        t5 = StringVar()
        t6 = StringVar()
        t7 = StringVar()
        t8 = StringVar()
        t9 = StringVar()
        t10 =StringVar()
        t11 =StringVar()
        t12 =StringVar()
        t13 =StringVar()
        t14 =StringVar()
        t15 =StringVar()
        t16 =StringVar()
        t17 =StringVar()
        t18 =StringVar()
        t19 =StringVar()

        def update(rows):
            trv.delete(*trv.get_children())
            for i in rows:
                trv.insert('', 'end', values=i)

        def search():
            q2 = q.get()
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            query = "SELECT id, date, name, company, dept, title, cell, tel, fax, email, country, state, " \
                    "city, street, postal, web, sns, gubun, memo FROM customer WHERE company " \
                    "LIKE '%" + q2 + "%' or name LIKE '%" + q2 + "%'"
            cursor.execute(query)
            rows = cursor.fetchall()
            update(rows)

        def clear():
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            query = "SELECT id, date, name, company, dept, title, cell, tel, fax, email, country, state, city, street, postal, web, sns, gubun, memo FROM customer"
            cursor.execute(query)
            rows = cursor.fetchall()
            update(rows)

        def getrow(event): #treeview에서 더블틀릭하면 입력란에 나타나게 하는 코드
            item = trv.item(trv.focus())
            values = item['values']
            if values:
                t1.set(values[0] if len(values) > 0 else '')
                t2.set(values[1] if len(values) > 1 else '')
                t3.set(values[2] if len(values) > 2 else '')
                t4.set(values[3] if len(values) > 3 else '')
                t5.set(values[4] if len(values) > 4 else '')
                t6.set(values[5] if len(values) > 5 else '')
                t7.set(values[6] if len(values) > 6 else '')
                t8.set(values[7] if len(values) > 7 else '')
                t9.set(values[8] if len(values) > 8 else '')
                t10.set(values[9] if len(values) > 9 else '')
                t11.set(values[10] if len(values) > 10 else '')
                t12.set(values[11] if len(values) > 11 else '')
                t13.set(values[12] if len(values) > 12 else '')
                t14.set(values[13] if len(values) > 13 else '')
                t15.set(values[14] if len(values) > 14 else '')
                t16.set(values[15] if len(values) > 15 else '')
                t17.set(values[16] if len(values) > 16 else '')
                t18.set(values[17] if len(values) > 17 else '')
                t19.set(values[18] if len(values) > 18 else '')

        def update_customer():
            fdate = t2.get()
            fname = t3.get()
            fcompany = t4.get()
            fdept = t5.get()
            ftitle = t6.get()
            fcell = t7.get()
            ftel = t8.get()
            ffax = t9.get()
            femail = t10.get()
            fcountry = t11.get()
            fstate = t12.get()
            fcity = t13.get()
            fstreet = t14.get()
            fpostal = t15.get()
            fweb = t16.get()
            fsns = t17.get()
            fgubun = t18.get()
            fmemo = t19.get()
            fid = t1.get()
            
            if messagebox.askyesno("Are You Sure you Want to update this custmer?"):
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                query = "UPDATE customer SET date = %s, name = %s, company = %s, dept = %s, title = %s, cell = %s, " \
                        "tel = %s, fax = %s, email = %s, country = %s, state = %s, city = %s, street = %s, postal = %s, " \
                        "web = %s, sns = %s, gubun = %s, memo = %s WHERE id = %s"
                cursor.execute(query, (fdate, fname, fcompany, fdept, ftitle, fcell, ftel, ffax, femail, fcountry, fstate, 
                                        fcity, fstreet, fpostal, fweb, fsns, fgubun, fmemo, fid))
                mydb.commit()
                clear()
            else:
                return True

        def add_new():
            fid = t1.get()
            fdate = t2.get()
            fname = t3.get()
            fcompany = t4.get()
            fdept = t5.get()
            ftitle = t6.get()
            fcell = t7.get()
            ftel = t8.get()
            ffax = t9.get()
            femail = t10.get()
            fcountry = t11.get()
            fstate = t12.get()
            fcity = t13.get()
            fstreet = t14.get()
            fpostal = t15.get()
            fweb = t16.get()
            fsns = t17.get()
            fgubun = t18.get()
            fmemo = t19.get()
            
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            # query = "INSERT INTO customer(SELECT date, name, company, dept, title, cell, tel, fax, email, country, state, " \
            #     "city, street, postal, web, sns, gubun, memo) VALUES( %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, %s, %s, %s,%s, %s,%s, %s,%s)"
            # cursor.execute(query, (fdate, fname, fcompany, fdept, ftitle, fcell, ftel, ffax, femail, fcountry, fstate, fcity, fstreet, fpostal, fweb, fsns, fgubun, fmemo))
            
            cursor.execute("insert into customer values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                                                        (t1.get(),
                                                                        t2.get(),
                                                                        t3.get(),
                                                                        t4.get(),
                                                                        t5.get(),
                                                                        t6.get(),
                                                                        t7.get(),
                                                                        t8.get(),
                                                                        t9.get(),
                                                                        t10.get(),
                                                                        t11.get(),
                                                                        t12.get(),
                                                                        t13.get(),
                                                                        t14.get(),
                                                                        t15.get(),
                                                                        t16.get(),
                                                                        t17.get(),
                                                                        t18.get(),
                                                                        t19.get()
                                                                        ))
            
            mydb.commit()
            clear()
            
            
        def delete_customer():
            fid = t1.get()
            if messagebox.askyesno("Confirm Delete?", "Are You Sure you Want to delete this customer?"):
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                query = "DELETE FROM customer WHERE id = "+ fid
                cursor.execute(query)
                mydb.commit()
                clear()
            else:
                return True

        def export_csv():
            return True

        def import_csv():
            return True

        def save_to_db():
            return True 
        
        #==========Search Section==================#
        lbl = Label(wrapper1, text="Search")
        lbl.pack(side=tk.LEFT, padx=10)
        ent = Entry(wrapper1, textvariable=q)
        ent.pack(side=tk.LEFT, padx=6)
        btn = Button(wrapper1, text="Search", command=search)
        btn.pack(side=tk.LEFT, padx=6)
        cbtn = Button(wrapper1, text="Clear", command=clear)
        cbtn.pack(side=tk.LEFT, padx=6)

        #===========User Data Input Section=========#
        lbl1 = tk.Label(wrapper2, text="ID")
        lbl1.grid(row=0, column=0, padx=5, pady=3)
        ent1 = tk.Entry(wrapper2, textvariable=t1)
        ent1.grid(row=1, column=0, padx=5, pady=3)

        lbl2 = tk.Label(wrapper2, text="Date")
        lbl2.grid(row=0, column=1, padx=5, pady=3)
        ent2 = tk.Entry(wrapper2, textvariable=t2)
        ent2.grid(row=1, column=1, padx=5, pady=3)

        lbl3 = tk.Label(wrapper2, text="Name")
        lbl3.grid(row=0, column=2, padx=5, pady=3)
        ent3 = tk.Entry(wrapper2, textvariable=t3)
        ent3.grid(row=1, column=2, padx=5, pady=3)

        lbl4 = tk.Label(wrapper2, text="Company")
        lbl4.grid(row=0, column=3, padx=5, pady=3)
        ent4 = tk.Entry(wrapper2, textvariable=t4)
        ent4.grid(row=1, column=3, padx=5, pady=3)

        lbl5 = tk.Label(wrapper2, text="Department")
        lbl5.grid(row=0, column=4, padx=5, pady=3)
        ent5 = tk.Entry(wrapper2, textvariable=t5)
        ent5.grid(row=1, column=4, padx=5, pady=3)

        lbl6 = tk.Label(wrapper2, text="Title")
        lbl6.grid(row=0, column=5, padx=5, pady=3)
        ent6 = tk.Entry(wrapper2, textvariable=t6)
        ent6.grid(row=1, column=5, padx=5, pady=3)

        lbl7 = tk.Label(wrapper2, text="Cell")
        lbl7.grid(row=0, column=6, padx=5, pady=3)
        ent7 = tk.Entry(wrapper2, textvariable=t7)
        ent7.grid(row=1, column=6, padx=5, pady=3)

        lbl8 = tk.Label(wrapper2, text="Tel")
        lbl8.grid(row=0, column=7, padx=5, pady=3)
        ent8 = tk.Entry(wrapper2, textvariable=t8)
        ent8.grid(row=1, column=7, padx=5, pady=3)

        lbl9 = tk.Label(wrapper2, text="Fax")
        lbl9.grid(row=0, column=8, padx=5, pady=3)
        ent9 = tk.Entry(wrapper2, textvariable=t9)
        ent9.grid(row=1, column=8, padx=5, pady=3)

        lbl10 = tk.Label(wrapper2, text="eMail")
        lbl10.grid(row=0, column=9, padx=5, pady=3)
        ent10 = tk.Entry(wrapper2, textvariable=t10)
        ent10.grid(row=1, column=9, padx=5, pady=3)

        lbl11 = tk.Label(wrapper2, text="Country")
        lbl11.grid(row=0, column=10, padx=5, pady=3)
        ent11 = tk.Entry(wrapper2, textvariable=t11)
        ent11.grid(row=1, column=10, padx=5, pady=3)

        lbl12 = tk.Label(wrapper2, text="State")
        lbl12.grid(row=0, column=11, padx=5, pady=3)
        ent12 = tk.Entry(wrapper2, textvariable=t12)
        ent12.grid(row=1, column=11, padx=5, pady=3)

        lbl13 = tk.Label(wrapper2, text="City")
        lbl13.grid(row=0, column=12, padx=5, pady=3)
        ent13 = tk.Entry(wrapper2, textvariable=t13)
        ent13.grid(row=1, column=12, padx=5, pady=3)

        lbl14 = tk.Label(wrapper2, text="Street")
        lbl14.grid(row=0, column=13, padx=5, pady=3)
        ent14 = tk.Entry(wrapper2, textvariable=t14)
        ent14.grid(row=1, column=13, padx=5, pady=3)

        lbl15 = tk.Label(wrapper2, text="Postal")
        lbl15.grid(row=0, column=14, padx=5, pady=3)
        ent15 = tk.Entry(wrapper2, textvariable=t15)
        ent15.grid(row=1, column=14, padx=5, pady=3)

        lbl16 = tk.Label(wrapper2, text="Web")
        lbl16.grid(row=0, column=15, padx=5, pady=3)
        ent16 = tk.Entry(wrapper2, textvariable=t16)
        ent16.grid(row=1, column=15, padx=5, pady=3)

        lbl17 = tk.Label(wrapper2, text="SNS")
        lbl17.grid(row=0, column=16, padx=5, pady=3)
        ent17 = tk.Entry(wrapper2, textvariable=t17)
        ent17.grid(row=1, column=16, padx=5, pady=3)

        lbl18 = tk.Label(wrapper2, text="Gubun")
        lbl18.grid(row=0, column=17, padx=5, pady=3)
        ent18 = tk.Entry(wrapper2, textvariable=t18)
        ent18.grid(row=1, column=17, padx=5, pady=3)

        lbl19 = tk.Label(wrapper2, text="Memo")
        lbl19.grid(row=0, column=18, padx=5, pady=3)
        ent19 = tk.Entry(wrapper2, textvariable=t19)
        ent19.grid(row=1, column=18, padx=5, pady=3)

        up_btn = tk.Button(wrapper2, text="Update", command=update_customer)
        add_btn = tk.Button(wrapper2, text="Add New", command=add_new)
        delete_btn = tk.Button(wrapper2, text="Delete", command=delete_customer)

        add_btn.grid(row=19, column=0, padx=5, pady=3)
        up_btn.grid(row=19, column=1, padx=5, pady=3)
        delete_btn.grid(row=19, column=2, padx=5, pady=3)

        #===========User Data list==================#
        
        trv = ttk.Treeview(wrapper3, columns=(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19), show="headings", height=6)
        trv.pack()

        trv.heading(1, text="ID")
        trv.heading(2, text="Date")
        trv.heading(3, text="Name")
        trv.heading(4, text="Company")
        trv.heading(5, text="Department")
        trv.heading(6, text="Title")
        trv.heading(7, text="Cell")
        trv.heading(8, text="Tel")
        trv.heading(9, text="Fax")
        trv.heading(10, text="eMail")
        trv.heading(11, text="Country")
        trv.heading(12, text="State")
        trv.heading(13, text="City")
        trv.heading(13, text="City")
        trv.heading(14, text="Street")
        trv.heading(15, text="Postal")
        trv.heading(16, text="Web")
        trv.heading(17, text="SNS")
        trv.heading(18, text="Gubun")
        trv.heading(19, text="Memo")
            
        trv.bind("<Double 1>", getrow) #<----------Error: NameError: name 'getrow' is not defined

        mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = mydb.cursor()
        query = "SELECT id, date, name, company, dept, title, cell, tel, fax, email, country, state, " \
            "city, street, postal, web, sns, gubun, memo FROM customer"
        cursor.execute(query)
        rows = cursor.fetchall()
        update(rows)

        expbtn = Button(wrapper3, text="Export to CSV", command="export_csv")
        expbtn.pack(side=LEFT, padx=10, pady=10)

        impbtn = Button(wrapper3, text="Import from CSV", command="import_csv")
        impbtn.pack(side=LEFT, padx=10, pady=10)

        savebtn = Button(wrapper3, text="Save to DB", command="save_to_db")
        savebtn.pack(side=LEFT, padx=10, pady=10)

        expbtn = Button(wrapper3, text="Exit", command=lambda: exit())
        expbtn.pack(side=LEFT, padx=10, pady=10)

#========== 여기까지 입니다==========#
class Employee(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="powder blue")
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Employee Registration System", bd=3, relief=RIDGE, 
                        bg="powder blue", fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)
        
        #connection for phpmyadmin
        def connection():
            conn = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            return conn

        def refreshTable():
            for data in my_tree.get_children():
                my_tree.delete(data)

            for array in read():
                my_tree.insert(parent='', index='end', iid=array, text="", values=(array), tag="orow")

            my_tree.tag_configure('orow', background='#EEEEEE', font=('Arial', 12))
            my_tree.grid(row=8, column=0, columnspan=5, rowspan=11, padx=10, pady=20)

        #placeholders for entry
        ph1 = tk.StringVar()
        ph2 = tk.StringVar()
        ph3 = tk.StringVar()
        ph4 = tk.StringVar()
        ph5 = tk.StringVar()
        ph6 = tk.StringVar()
        ph7 = tk.StringVar()
        ph8 = tk.StringVar()
        ph9 = tk.StringVar()
        ph10 = tk.StringVar()

        #placeholder set value function
        def setph(word,num):
            if num ==1:
                ph1.set(word)
            if num ==2:
                ph2.set(word)
            if num ==3:
                ph3.set(word)
            if num ==4:
                ph4.set(word)
            if num ==5:
                ph5.set(word)
            if num ==6:
                ph6.set(word)
            if num ==7:
                ph7.set(word)
            if num ==8:
                ph8.set(word)
            if num ==9:
                ph9.set(word)
            if num ==10:
                ph10.set(word)

        def read():
            conn = connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM employees")
            results = cursor.fetchall()
            conn.commit()
            conn.close()
            return results

        def add():
            fid = str(idEntry.get())
            fname = str(nameEntry.get())
            frole = str(roleEntry.get())
            fgender = str(genderEntry.get())
            fstatus = str(statusEntry.get())
            femail = str(emailEntry.get())
            fcell = str(cellEntry.get())
            faddress = str(addressEntry.get())
            fpost = str(postEntry.get())
            fsalary = salaryEntry.get()

            if not fid.strip() or not fname.strip() or not frole.strip() or not fgender.strip() or not fstatus.strip():
                messagebox.showinfo("Error", "빈 칸을 채워주세요.")
                return
            else:
                try:
                    if fsalary == "":  # 빈 문자열인 경우 0으로 설정합니다.
                        fsalary = 0
                    else:
                        fsalary = int(fsalary) # 정수로 변환합니다.
                    conn = connection()
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO employees VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                (fid, fname, frole, fgender, fstatus, femail, fcell, faddress, fpost, fsalary))
                    conn.commit()
                    conn.close()
                    messagebox.showinfo("Success", "직원 정보가 추가되었습니다.")
                except pymysql.IntegrityError:
                    messagebox.showinfo("Error", "이미 존재하는 직원 ID입니다.")
                    return

            refreshTable()

        def reset():
            decision = messagebox.askquestion("Warning!!", "Delete all data?")
            if decision != "yes":
                return 
            else:
                try:
                    conn = connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM employees")
                    conn.commit()
                    conn.close()
                except:
                    messagebox.showinfo("Error", "Sorry an error occured")
                    return

                refreshTable()

        def delete():
            decision = messagebox.askquestion("Warning!!", "Delete the selected data?")
            if decision != "yes":
                return 
            else:
                selected_item = my_tree.selection()[0]
                deleteData = str(my_tree.item(selected_item)['values'][0])
                try:
                    conn = connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM employees WHERE id='"+str(deleteData)+"'")
                    conn.commit()
                    conn.close()
                except:
                    messagebox.showinfo("Error", "Sorry an error occured")
                    return

                refreshTable()

        def select():
            try:
                selected_item = my_tree.selection()[0]
                fid = str(my_tree.item(selected_item)['values'][0])
                fname = str(my_tree.item(selected_item)['values'][1])
                frole = str(my_tree.item(selected_item)['values'][2])
                fgender = str(my_tree.item(selected_item)['values'][3])
                fstatus = str(my_tree.item(selected_item)['values'][4])
                femail = str(my_tree.item(selected_item)['values'][5])
                fcell = str(my_tree.item(selected_item)['values'][6])
                faddress = str(my_tree.item(selected_item)['values'][7])
                fpost = str(my_tree.item(selected_item)['values'][8])
                fsalary = str(my_tree.item(selected_item)['values'][9])

                setph(fid,1)
                setph(fname,2)
                setph(frole,3)
                setph(fgender,4)
                setph(fstatus,5)
                setph(femail, 6)
                setph(fcell, 7)
                setph(faddress, 8)
                setph(fpost, 9)
                setph(fsalary, 10)
                
            except:
                messagebox.showinfo("Error", "Please select a data row")

        def search():
            fid = str(idEntry.get())
            fname = str(nameEntry.get())
            frole = str(roleEntry.get())
            fgender = str(genderEntry.get())
            fstatus = str(statusEntry.get())
            femail = str(emailEntry.get())
            fcell = str(cellEntry.get())
            faddress = str(addressEntry.get())
            fpost = str(postEntry.get())
            fsalary = str(salaryEntry.get())

            conn = connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM employees WHERE id='"+
            fid+"' or id='"+
            fname+"' or name='"+
            frole+"' or role='"+
            fgender+"' or gender='"+
            fstatus+"' or status='"+
            femail+"' or email='"+
            fcell+"' or cell='"+
            faddress+"' or address='"+
            fpost+"' or post='"+
            fsalary+"' or salary='"+
            fsalary+"' ")
            
            try:
                result = cursor.fetchall()

                for num in range(0,10):
                    setph(result[0][num],(num+1))

                conn.commit()
                conn.close()
            except:
                messagebox.showinfo("Error", "No data found")

        def update():
            selectedfid = ""

            try:
                selected_item = my_tree.selection()[0]
                selectedid = str(my_tree.item(selected_item)['values'][0])
            except:
                messagebox.showinfo("Error", "Please select a data row")

            fid = str(idEntry.get())
            fname = str(nameEntry.get())
            frole = str(roleEntry.get())
            fgender = str(genderEntry.get())
            fstatus = str(statusEntry.get())
            femail = str(emailEntry.get())
            fcell = str(cellEntry.get())
            faddress = str(addressEntry.get())
            fpost = str(postEntry.get())
            fsalary = str(salaryEntry.get())

            if (fid == "" or fid == " ") or (fname == "" or fname == " ") or (frole == "" or frole == " ") or (fgender == "" or fgender == " ") or (fstatus == "" or fstatus == " "):
                messagebox.showinfo("Error", "Please fill up the blank entry")
                return
            else:
                try:
                    conn = connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE employees SET id='"+
                    fid+"', name='"+
                    fname+"', role='"+
                    frole+"', gender='"+
                    fgender+"', status='"+
                    fstatus+"', email='"+
                    femail+"', cell='"+
                    fcell+"', address='"+
                    faddress+"', post='"+
                    fpost+"', salary='"+
                    fsalary+"' WHERE id='"+
                    selectedid+"' ")
                    conn.commit()
                    conn.close()
                except:
                    messagebox.showinfo("Error", "ID already exist")
                    return

            refreshTable()

        emp_wrapper1 = tk.LabelFrame(self, text="Employee Infomation Input", font=('Arial', 12), bg="powder blue")
        emp_wrapper1.place(x=6, y=50, width=1800, height=170)

        idLabel = Label(emp_wrapper1, text="ID", font=('Arial', 11))
        nameLabel = Label(emp_wrapper1, text="Name", font=('Arial', 11))
        rolLabel = Label(emp_wrapper1, text="Role", font=('Arial', 11))
        genderLabel = Label(emp_wrapper1, text="Gender", font=('Arial', 11))
        statusLabel = Label(emp_wrapper1, text="Status", font=('Arial', 11))
        emailLabel = Label(emp_wrapper1, text="Email", font=('Arial', 11))
        cellLabel = Label(emp_wrapper1, text="Cell", font=('Arial', 11))
        addressLabel = Label(emp_wrapper1, text="Address", font=('Arial', 11))
        postLabel = Label(emp_wrapper1, text="Post", font=('Arial', 11))
        salaryLabel = Label(emp_wrapper1, text="Salary", font=('Arial', 11))

        idLabel.grid(row=3, column=0, columnspan=1, padx=(5,2), pady=5, sticky="w")
        nameLabel.grid(row=3, column=1, columnspan=1, padx=5, pady=5, sticky="w")
        rolLabel.grid(row=3, column=2, columnspan=1, padx=5, pady=5, sticky="w")
        genderLabel.grid(row=3, column=3, columnspan=1, padx=5, pady=5, sticky="w")
        statusLabel.grid(row=3, column=4, columnspan=1, padx=5, pady=5, sticky="w")
        emailLabel.grid(row=3, column=5, columnspan=1, padx=5, pady=5, sticky="w")
        cellLabel.grid(row=3, column=6, columnspan=1, padx=5, pady=5, sticky="w")
        addressLabel.grid(row=3, column=7, columnspan=1, padx=5, pady=5, sticky="w")
        postLabel.grid(row=3, column=8, columnspan=1, padx=5, pady=5, sticky="w")
        salaryLabel.grid(row=3, column=9, columnspan=1, padx=5, pady=5, sticky="w")

        idEntry = Entry(emp_wrapper1, width=5, bd=5, font=('Arial', 11), textvariable = ph1)
        nameEntry = Entry(emp_wrapper1, width=20, bd=5, font=('Arial', 11), textvariable = ph2)
        roleEntry = Entry(emp_wrapper1, width=15, bd=5, font=('Arial', 11), textvariable = ph3)
        genderEntry = Entry(emp_wrapper1, width=10, bd=5, font=('Arial', 11), textvariable = ph4)
        statusEntry = Entry(emp_wrapper1, width=7, bd=5, font=('Arial', 11), textvariable = ph5)
        emailEntry = Entry(emp_wrapper1, width=30, bd=5, font=('Arial', 11), textvariable = ph6)
        cellEntry = Entry(emp_wrapper1, width=10, bd=5, font=('Arial', 11), textvariable = ph7)
        addressEntry = Entry(emp_wrapper1, width=50, bd=5, font=('Arial', 11), textvariable = ph8)
        postEntry = Entry(emp_wrapper1, width=10, bd=1, font=('Arial', 11), textvariable = ph9)
        salaryEntry = Entry(emp_wrapper1, width=13, bd=5, font=('Arial', 11), textvariable = ph10)

        idEntry.grid(row=4, column=0, columnspan=1, padx=(5,2), pady=5, sticky="w")
        nameEntry.grid(row=4, column=1, columnspan=1, padx=5, pady=5, sticky="w")
        roleEntry.grid(row=4, column=2, columnspan=1, padx=5, pady=5, sticky="w")
        genderEntry.grid(row=4, column=3, columnspan=1, padx=5, pady=5, sticky="w")
        statusEntry.grid(row=4, column=4, columnspan=1, padx=5, pady=5, sticky="w")
        emailEntry.grid(row=4, column=5, columnspan=1, padx=5, pady=5, sticky="w")
        cellEntry.grid(row=4, column=6, columnspan=1, padx=5, pady=5, sticky="w")
        addressEntry.grid(row=4, column=7, columnspan=1, padx=5, pady=5, sticky="w")
        postEntry.grid(row=4, column=8, columnspan=1, padx=5, pady=5, sticky="w")
        salaryEntry.grid(row=4, column=9, columnspan=1, padx=5, pady=5, sticky="w")

        addBtn = Button(emp_wrapper1, text="Add", padx=10, pady=10, width=5, height=1, bd=5, font=('Arial', 11), bg="#84F894", command=add)
        updateBtn = Button(emp_wrapper1, text="Update", padx=10, pady=10, width=5, height=1, bd=5, font=('Arial', 11), bg="#84E8F8", command=update)
        deleteBtn = Button(emp_wrapper1, text="Delete", padx=10, pady=10, width=5, height=1, bd=5, font=('Arial', 11), bg="#FF9999", command=delete)
        searchBtn = Button(emp_wrapper1, text="Search", padx=10, pady=10, width=5, height=1, bd=5, font=('Arial', 11), bg="#F4FE82", command=search)
        resetBtn = Button(emp_wrapper1, text="Reset", padx=10, pady=10, width=5, height=1,   bd=5, font=('Arial', 11), bg="#F398FF", command=reset)
        selectBtn = Button(emp_wrapper1, text="Select", padx=10, pady=10, width=1, height=1, bd=5, font=('Arial', 11), bg="#EEEEEE", command=select)

        addBtn.grid(row=6, column=0, padx=5, pady=10, sticky="ew")
        updateBtn.grid(row=6, column=1, padx=5, pady=10, sticky="ew")
        deleteBtn.grid(row=6, column=2, padx=5, pady=10, sticky="ew")
        searchBtn.grid(row=6, column=3, padx=5, pady=10, sticky="ew")
        resetBtn.grid(row=6, column=4, padx=5, pady=10, sticky="ew")
        selectBtn.grid(row=6, column=5, padx=5, pady=10, sticky="ew")

        #======================================
        emp_wrapper2 = tk.LabelFrame(self, text="Employee Data List", font=('Arial', 12), bg="powder blue")
        emp_wrapper2.place(x=6, y=237, width=1800, height=610)
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 10))

        my_tree = ttk.Treeview(emp_wrapper2, columns=("id","name","role","gender","status","email","cell","address","post","salary"), show="headings", height=6)
        my_tree.grid()

        my_tree.column("#0", width=0, stretch=tk.NO)
        my_tree.column("id", anchor=tk.W, width=100)
        my_tree.column("name", anchor=tk.W, width=150)
        my_tree.column("role", anchor=tk.W, width=100)
        my_tree.column("gender", anchor=tk.W, width=50)
        my_tree.column("status", anchor=tk.W, width=50)
        my_tree.column("email", anchor=tk.W, width=200)
        my_tree.column("cell", anchor=tk.W, width=150)
        my_tree.column("address", anchor=tk.W, width=300)
        my_tree.column("post", anchor=tk.W, width=100)
        my_tree.column("salary", anchor=tk.W, width=100)

        my_tree.heading("id", text="ID", anchor=tk.W)
        my_tree.heading("name", text="Name", anchor=tk.W)
        my_tree.heading("role", text="Role", anchor=tk.W)
        my_tree.heading("gender", text="Gender", anchor=tk.W)
        my_tree.heading("status", text="Status", anchor=tk.W)
        my_tree.heading("email", text="Email", anchor=tk.W)
        my_tree.heading("cell", text="Cell", anchor=tk.W)
        my_tree.heading("address", text="Address", anchor=tk.W)
        my_tree.heading("post", text="Post", anchor=tk.W)
        my_tree.heading("salary", text="Salary", anchor=tk.W)

        refreshTable()
        
#========== 여기부터 코드 작성 ==========#
class ImageSaveDB(tk.Frame):  
    def __init__(self, parent):
        super().__init__(parent, bg='powder blue')
        self.pack(fill=BOTH, expand=True)
        label = Label(self, text="Image File to Save on MySQL_DB")
        label.pack()
        
        def insert_blob(file_path, gubun):
            with open(file_path, "rb") as file:
                binary_data = file.read()
            file_name = os.path.basename(file_path)
            sql_statement = "INSERT INTO receiptimages (Photo, filename, gubun) VALUES (%s, %s, %s)"
            cursor.execute(sql_statement, (binary_data, file_name, gubun))
            myDB.commit()


        from PIL import Image

        def retrieve_blob(selected_id):
            sql_statement = "SELECT filename, gubun FROM receiptimages WHERE id = %s"
            cursor.execute(sql_statement, (selected_id,))
            result = cursor.fetchone()
            if result is not None:
                binary_data, file_name = result
                image = Image.open(io.BytesIO(binary_data))
                photo = ImageTk.PhotoImage(image)
                image_label.config(image=photo)
                image_label.image = photo
                save_path_var.set(os.path.join(selected_save_path_var.get(), file_name))

        def browse_file():
            file_path = filedialog.askopenfilename()
            selected_image_path_var.set(file_path)

        def browse_save_path():
            save_path = filedialog.askdirectory()
            selected_save_path_var.set(save_path)
            save_path_var.set(save_path)

        def save_file():
            if not selected_save_path_var.get():
                messagebox.showerror("오류", "저장 경로를 지정해주세요.")
                return
            insert_blob(selected_image_path_var.get(), selected_gubun_var.get())
            refresh_image_list()

        def refresh_image_list():
            # sql_statement = "SELECT id, filename FROM receiptimages WHERE gubun = %s"
            # cursor.execute(sql_statement, (selected_gubun_var.get(),))
            # image_list = cursor.fetchall()
            # for item in image_tree.get_children():
            #     image_tree.delete(item)
            # for id, file_name in image_list:
            #     image_tree.insert("", "end", values=(id, file_name))
            
            sql_statement = "SELECT * FROM receiptimages"
            cursor.execute(sql_statement)
            image_list = cursor.fetchall()
            for item in image_tree.get_children():
                image_tree.delete(item)
            for row in image_list:
                image_tree.insert("", "end", values=row)

        def on_image_select(event):
            item = image_tree.selection()[0]
            selected_id = image_tree.item(item, "values")[0]
            retrieve_blob(selected_id)

        myDB = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = myDB.cursor()

        selected_image_path_var = tk.StringVar()
        selected_save_path_var = tk.StringVar()
        selected_gubun_var = tk.StringVar()
        save_path_var = tk.StringVar()

        image_frame = tk.Frame(self)
        image_frame.pack(side=tk.LEFT, padx=10, pady=10)

        image_label = tk.Label(image_frame)
        image_label.pack()

        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.LEFT, padx=10, pady=10)

        image_path_label = tk.Label(button_frame, text="이미지 경로:")
        image_path_label.grid(row=0, column=0, padx=5, pady=5)

        image_path_entry = tk.Entry(button_frame, textvariable=selected_image_path_var, width=30)
        image_path_entry.grid(row=0, column=1, padx=5, pady=5)

        browse_button = tk.Button(button_frame, text="찾아보기", command=browse_file)
        browse_button.grid(row=0, column=2, padx=5, pady=5)

        save_path_label = tk.Label(button_frame, text="저장 경로:")
        save_path_label.grid(row=1, column=0, padx=5, pady=5)

        save_path_entry = tk.Entry(button_frame, textvariable=save_path_var, width=30)
        save_path_entry.grid(row=1, column=1, padx=5, pady=5)

        browse_save_path_button = tk.Button(button_frame, text="저장 경로 찾기", command=browse_save_path)
        browse_save_path_button.grid(row=1, column=2, padx=5, pady=5)

        gubun_label = tk.Label(button_frame, text="구분:")
        gubun_label.grid(row=2, column=0, padx=5, pady=5)

        gubun_combo = ttk.Combobox(button_frame, textvariable=selected_gubun_var)
        gubun_combo['values'] = ('자산구매', '자산처분', '감가상각', '부채발생', '부채감소', '대표차입금','세금게산서','상품구매','원가부대비용','비용','수익증가','수익감소','기타')
        gubun_combo.grid(row=2, column=1, padx=5, pady=5)

        save_button = tk.Button(button_frame, text="저장", command=save_file)
        save_button.grid(row=3, columnspan=3, padx=5, pady=5)

        image_tree_frame = tk.Frame(self)
        image_tree_frame.pack(side=tk.LEFT, padx=10, pady=10)

        image_tree = ttk.Treeview(image_tree_frame, columns=("id", "filename", "gubun", "photo"), displaycolumns=("id", "filename", "gubun", "photo"))
        image_tree.heading("#0", text="ID")
        image_tree.heading("#1", text="File Name")
        image_tree.heading("#2", text="Gubun")
        image_tree.heading("#3", text="Photo")
        image_tree.pack()

        image_tree.bind("<<TreeviewSelect>>", on_image_select)

        refresh_image_list()
        
#========== 여기부터 코드 작성 ==========#
class Aledger(tk.Frame):  
    def __init__(self, parent):
        super().__init__(parent, bg='#161C25')
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Other Transaction Input System", bd=3, relief=RIDGE, 
                        bg="powder blue", fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)

        font1 = ('Arial', 12, 'bold')
        font2 = ('Arial', 9)
        today_date = datetime.date.today().strftime('%Y-%m-%d')

        # -----Variable declare and the maximum value of the aid field from the aledger table
        mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
        cursor=mydb.cursor()
        cursor.execute("SELECT MAX(aid) FROM aledger")

        result = cursor.fetchone()
        if result is not None:
            max_aid = result[0]
        else:
            max_aid = 0  # Assign a default value of 0 if max_tsq is None

        aid_var = StringVar()
        aid_var.set(str(max_aid + 1))
        linesq_var = StringVar()
        linesq_var.set(str(aid_var.get()) + "-1")


        # 이전에 'float' 값을 설정한 코드를 수정하여 'StringVar'로 변환합니다.


        # MySQL 데이터베이스 연결
        def fetch_aledger():
            try:
                # MySQL 데이터베이스 연결
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                cursor.execute("SELECT * FROM aledger")
                result = cursor.fetchall()
                return result
            except pymysql.Error as e:
                print("MySQL Error:", e)
            finally: # 연결 종료
                cursor.close()
                mydb.close()

        def fetch_coa_values():
            try:
                # MySQL 데이터베이스 연결
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                cursor.execute("SELECT COA2 FROM coa")
                result = cursor.fetchall()
                coa_values = [row[0] for row in result]  # 튜플에서 값을 추출하여 리스트로 변환
                return coa_values
            except pymysql.Error as e:
                print("MySQL Error:", e)
            finally:  # 연결 종료
                if 'cursor' in locals():
                    cursor.close()
                if 'mydb' in locals():
                    mydb.close()

        # def _textvariable_callback(self, *_):
        #     value = self._textvariable.get()
        #     if value == "":
        #         self._value = 0.0
        #     else:
        #         try:
        #             self._value = self._textvariable.get()  # DoubleVar이므로 그대로 가져옵니다.
        #         except ValueError:
        #             self._value = 0.0
                    
        def calculate_net(*args):
            try:
                total = float(total_entry.get())
                vat = float(vat_entry.get())
                net = total-vat
                net_var.set(str(net))
                # 순수익 라벨 업데이트
                #self.lbl_net.config(text="{:.2f}".format(due))
            except ValueError:
                pass 

        def add_to_treeview():
            rows = fetch_aledger()
            if rows is not None:
                tree.delete(*tree.get_children())
                for row in rows:
                    tree.insert('', tk.END, values=row)

        def clear(*clicked):
            if clicked:
                tree.selection_remove(tree.focus())
                tree.focus('')
            aid_entry.delete(0, END)
            indate_entry.delete(0, END)
            gubun_combobox.set('')
            linesq_entry.delete(0, END)
            voudate_entry.delete(0, END)
            supplier_entry.delete(0, END)
            coa_options.set('')
            #Variable1.set('Male')
            total_entry.delete(0, END)
            vat_entry.delete(0, END)
            net_entry.delete(0, END)
            descript_entry.delete(0, END)
            currency_entry.delete(0, END)
            vouno_entry.delete(0, END)
            memo_entry.delete(0, END)
            agroup_combobox.set('')
            
        def display_data(event):
            selected_item = tree.focus()
            if selected_item:
                row = tree.item(selected_item)['values']
                clear()
                aid_entry.insert(0, row[0])
                indate_entry.insert(0, row[1])
                gubun_combobox.set(row[2])
                linesq_entry.insert(0, row[3])
                voudate_entry.insert(0, row[4])
                supplier_entry.insert(0, row[5])
                coa_options.set(row[6])
                total_entry.insert(0, row[7])
                vat_entry.insert(0, row[8])
                net_entry.delete(0, END)
                net_entry.insert(0, row[9])
                descript_entry.insert(0, row[10])
                currency_entry.insert(0, row[11])
                vouno_entry.insert(0, row[12])
                memo_entry.insert(0, row[13])
                agroup_combobox.set(row[14])
                
            else:
                pass


        def delete():
            selected_item = tree.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to delete.')
            else:
                aid = aid_entry.get()
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "DELETE from aledger where aid=%s"
                    val = (aid,) #'val' 변수가 튜플이 아닌 함수로 설정되어 있기 때문입니다. 'val,' 변수를 튜플로 설정해야 합니다.
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    add_to_treeview()
                    clear()
                    messagebox.showinfo('Success', 'Data has been deleted.')

                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()

        def update():
            selected_item = tree.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to update.')
            else:
                aid = aid_entry.get()
                indate = indate_entry.get()
                gubun = gubun_combobox.get()
                linesq = linesq_entry.get()
                voudate = voudate_entry.get()
                supplier = supplier_entry.get()
                coa = coa_options.get()
                total = total_entry.get()
                vat = vat_entry.get()
                net = net_entry.get()
                descript = descript_entry.get()
                currency = currency_entry.get()
                vouno = vouno_entry.get()
                memo = memo_entry.get()
                agroup = agroup_combobox.get()
                
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "UPDATE aledger set indate=%s,gubun=%s,linesq=%s,voudate=%s,supplier=%s,coa=%s,total=%s,vat=%s,net=%s,descript=%s," \
                            "currency=%s,vouno=%s,memo=%s,agroup=%s where aid=%s"
                    val = (indate, gubun, linesq, voudate, supplier, coa, total, vat, net, descript, currency, vouno, memo, agroup,aid)
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    add_to_treeview()
                    messagebox.showinfo('Success', 'Data has been updated.')
                
                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not updated successfully...")    
                

        def insert():
            #aid = aid_entry.get()
            indate = indate_entry.get()
            gubun = gubun_combobox.get()
            linesq = linesq_entry.get()
            voudate = voudate_entry.get()
            supplier = supplier_entry.get()
            coa = coa_options.get()
            total = total_entry.get()
            vat = vat_entry.get()
            net = net_entry.get()
            descript = descript_entry.get()
            currency = currency_entry.get()
            vouno = vouno_entry.get()
            memo = memo_entry.get()
            agroup = agroup_combobox.get()
            
            mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
            cursor=mydb.cursor()
            
            if not (id and voudate and gubun and coa and net and vouno and agroup):
                messagebox.showerror('Error', 'Enter all fields.')
            else:
                try:
                    sql = "INSERT INTO  aledger (indate, gubun, linesq, voudate, supplier, coa, total, vat, net, descript, currency, vouno, memo, agroup)" \
                        "VALUES (%s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                    val = (indate, gubun, linesq, voudate, supplier, coa, total, vat, net, descript, currency, vouno, memo, agroup)
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    messagebox.showinfo('Success', 'Data has been inserted.')
                
                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not inserted successfully...")    
                add_to_treeview()

        #---------Variable 선언                    
        coa_values = fetch_coa_values()
        Variable1 = StringVar()
        option2 = ['자산구매', '자산처분', '감가상각', '부채발생', '부채감소', '대표대여금', '세금계산서', '상품구매','기타원가항목','기타']
        Variable2 = StringVar()
        option3 = ['CASH', 'BANK', 'CreditCard', 'A/P', 'ACCRUD','ETC']
        Variable3 = StringVar()

        total_var = StringVar()
        vat_var = StringVar()
        net_var = StringVar()

        #===========upper Frame=======================
        mainframe = Frame(parent, bd=3, relief=RIDGE,padx=10, bg="#161C25")
        mainframe.place(x=0, y=50, width=1900, height=1000)
        
        #=====Sales_Header_Input====
        leftframe=LabelFrame(mainframe, text="All Transaction Input Except Receipts", bg="#161C25",bd=2, relief=RIDGE,font=font1,padx=1, fg='white')
        leftframe.place(x=0, y=2, width=300, height=700)

        #---------Labels and Entries 
        aid_label = customtkinter.CTkLabel(leftframe, font=font1, text='ID:', text_color='#fff', bg_color='#161C25')
        indate_label = customtkinter.CTkLabel(leftframe, font=font1, text='InDate:', text_color='#fff', bg_color='#161C25')
        gubun_label = customtkinter.CTkLabel(leftframe, font=font1, text='Gubun:', text_color='#fff', bg_color='#161C25')
        linesq_label = customtkinter.CTkLabel(leftframe, font=font1, text='LineSQ:', text_color='#fff', bg_color='#161C25')
        voudate_label = customtkinter.CTkLabel(leftframe, font=font1, text='VouDate:', text_color='#fff', bg_color='#161C25')
        supplier_label = customtkinter.CTkLabel(leftframe, font=font1, text='Supplier:', text_color='#fff', bg_color='#161C25')
        coa_label = customtkinter.CTkLabel(leftframe, font=font1, text='COA:', text_color='#fff', bg_color='#161C25')
        total_label = customtkinter.CTkLabel(leftframe, font=font1, text='Total:', text_color='#fff', bg_color='#161C25')
        vat_label = customtkinter.CTkLabel(leftframe, font=font1, text='VAT:', text_color='#fff', bg_color='#161C25')
        net_label = customtkinter.CTkLabel(leftframe, font=font1, text='Net:', text_color='#fff', bg_color='#161C25') 
        descript_label = customtkinter.CTkLabel(leftframe, font=font1, text='Description:', text_color='#fff', bg_color='#161C25')
        currency_label = customtkinter.CTkLabel(leftframe, font=font1, text='Currency:', text_color='#fff', bg_color='#161C25')
        vouno_label = customtkinter.CTkLabel(leftframe, font=font1, text='Voucher No:', text_color='#fff', bg_color='#161C25')    
        memo_label = customtkinter.CTkLabel(leftframe, font=font1, text='Memo:', text_color='#fff', bg_color='#161C25')
        agroup_label = customtkinter.CTkLabel(leftframe, font=font1, text='Tras Group:', text_color='#fff', bg_color='#161C25')

        aid_label.place(x=20, y=20)
        indate_label.place(x=20, y=55)
        gubun_label.place(x=20, y=90)
        linesq_label.place(x=20, y=125)
        voudate_label.place(x=20, y=160)
        supplier_label.place(x=20, y=195)
        coa_label.place(x=20, y=230)
        total_label.place(x=20, y=265)
        vat_label.place(x=20, y=300)
        net_label.place(x=20, y=335)
        descript_label.place(x=20, y=370)
        currency_label.place(x=20, y=405)
        vouno_label.place(x=20, y=440)
        memo_label.place(x=20, y=475)
        agroup_label.place(x=20, y=510)

        aid_entry = customtkinter.CTkEntry(leftframe, font=font1, textvariable=aid_var,text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        indate_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        indate_entry.insert(0, today_date)  # 기본값으로 오늘 날짜 설정
        gubun_combobox = customtkinter.CTkComboBox(leftframe, font=font1, text_color='#000', fg_color='#fff', dropdown_hover_color='#0C9295', button_color='#0C9295',
                                                    button_hover_color='#0C9295', border_color='#0C9295', width=180, variable=Variable3, values=option3,state='readonly')
        linesq_entry = customtkinter.CTkEntry(leftframe, font=font1, textvariable=linesq_var,text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        voudate_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        voudate_entry.insert(0, today_date)  # 기본값으로 오늘 날짜 설정
        supplier_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        coa_options = customtkinter.CTkComboBox(leftframe, font=font1, text_color='#000', fg_color='#fff', dropdown_hover_color='#0C9295', button_color='#0C9295', 
                                                button_hover_color='#0C9295', border_color='#0C9295', width=180, variable=Variable1, values=coa_values,state='readonly')
        total_entry = customtkinter.CTkEntry(leftframe, font=font1, textvariable=total_var, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        vat_entry = customtkinter.CTkEntry(leftframe, font=font1, textvariable=vat_var, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        net_entry = customtkinter.CTkEntry(leftframe, font=font1, textvariable=net_var, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        descript_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        currency_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        vouno_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        memo_entry = customtkinter.CTkEntry(leftframe, font=font1, text_color='#000', fg_color='#fff', border_color='#0C9295', border_width=2, width=180)
        agroup_combobox = customtkinter.CTkComboBox(leftframe, font=font1, text_color='#000', fg_color='#fff', dropdown_hover_color='#0C9295', button_color='#0C9295',
                                                    button_hover_color='#0C9295', border_color='#0C9295', width=180, variable=Variable2, values=option2,state='readonly')

        aid_entry.place(x=110, y=20)
        indate_entry.place(x=110, y=55)
        gubun_combobox.place(x=110,y=90)
        linesq_entry.place(x=110, y=125)
        voudate_entry.place(x=110, y=160)
        supplier_entry.place(x=110, y=195)
        coa_options.place(x=110, y=230)
        total_entry.place(x=110, y=265)
        vat_entry.place(x=110, y=300)
        net_entry.place(x=110, y=335)
        descript_entry.place(x=110, y=370)
        currency_entry.place(x=110, y=405)
        vouno_entry.place(x=110, y=440)
        memo_entry.place(x=110, y=475)
        agroup_combobox.place(x=110, y=510)

        # calculate_net 메소드를 호출하여 var_net 값을 계산합니다.
        # trace() 메소드를 사용하여 var_total과 var_vat 값이 변경될 때마다 calculate_net() 메소드를 호출하도록 설정합니다.
        total_var.trace('w', calculate_net)
        total_var.set(str(0.0))  # Convert the value to float
        vat_var.trace('w', calculate_net)
        vat_var.set(str(0.0))  # Convert the value to float

        #---------Buttons
        add_button = customtkinter.CTkButton(mainframe, command=insert, font=font1, text='Add Transaction', text_color='#fff', 
                                            fg_color='#05A312', hover_color='#00850B', bg_color='#161C25', cursor='hand2', corner_radius=15, width=260)
        add_button.place(x=20, y=740)

        clear_button = customtkinter.CTkButton(mainframe, command=lambda:clear(True), font=font1, text='New Transaction', text_color='#fff', fg_color='#161C25', 
                                                hover_color='#FF5002', bg_color='#161C25', border_color='#F15704', border_width=2, cursor='hand2', corner_radius=15, width=260)
        clear_button.place(x=20,y=770)

        update_button = customtkinter.CTkButton(mainframe, command=update, font=font1, text='Update Transaction', text_color='#fff', fg_color='#161C25', 
                                                hover_color='#FF5002', bg_color='#161C25', border_color='#F15704', border_width=2, cursor='hand2', corner_radius=15, width=260)
        update_button.place(x=300, y=770)

        delete_button = customtkinter.CTkButton(mainframe, command=delete, font=font1, text='Delete Transaction', text_color='#fff', fg_color='#161C25', 
                                                hover_color='#AE0000', bg_color='#161C25', border_color='#E40404', border_width=2, cursor='hand2', corner_radius=15, width=260)
        delete_button.place(x=580, y=770)

        #=====Sales_Header_Input====
        rightframe=LabelFrame(mainframe, text="Other Transaction DB", bg="#161C25",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=2, fg='white')
        rightframe.place(x=300, y=1, width=1500, height=700)
        
        #---------Treeview
        style = ttk.Style(rightframe)
        style.theme_use('clam')
        style.configure('Treeview', font=font2, foreground='#fff', background='#000', fieldbackground='#313837')
        style.map('Treeview', background=[('selected', '#1A8F2D')])

        tree = ttk.Treeview(rightframe, height=300)
        tree['columns'] = ('aid', 'indate', 'gubun', 'linesq', 'voudate', 'supplier', 'coa', 'total', 'vat', 'net', 'descript', 'currency', 'vouno', 'memo', 'agroup')

        tree.column('#0', width=0, stretch=NO) #Hide the default column
        tree.column('aid', anchor=CENTER, width=50)
        tree.column('indate', anchor=CENTER, width=100)
        tree.column('gubun', anchor=CENTER, width=80)
        tree.column('linesq', anchor=CENTER, width=60)
        tree.column('voudate', anchor=CENTER, width=100)
        tree.column('supplier', anchor=CENTER, width=120)
        tree.column('coa', anchor=CENTER, width=120)
        tree.column('total', anchor=CENTER, width=100)
        tree.column('vat', anchor=CENTER, width=100)
        tree.column('net', anchor=CENTER, width=100)
        tree.column('descript', anchor=CENTER, width=120)
        tree.column('currency', anchor=CENTER, width=50)
        tree.column('vouno', anchor=CENTER, width=120)
        tree.column('memo', anchor=CENTER, width=120)
        tree.column('agroup', anchor=CENTER, width=120)

        tree.heading('aid', text='ID')
        tree.heading('indate', text='Date')
        tree.heading('gubun', text='Type')
        tree.heading('linesq', text='Line')
        tree.heading('voudate', text='Voucher Date')
        tree.heading('supplier', text='Supplier')
        tree.heading('coa', text='COA')
        tree.heading('total', text='Total')
        tree.heading('vat', text='VAT')
        tree.heading('net', text='Net')
        tree.heading('descript', text='Description')
        tree.heading('currency', text='Currency')
        tree.heading('vouno', text='Voucher No')
        tree.heading('memo', text='Memo')
        tree.heading('agroup', text='Account Group')

        tree.place(x=2, y=7)

        tree.bind('<ButtonRelease>', display_data)

        add_to_treeview()

#========== 여기부터 코드 작성 ==========#
class Invoice(tk.Frame):  
    def __init__(self, parent):
        super().__init__(parent, bg="powder blue")
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Sales Invoice Management System", bd=3, relief=RIDGE, 
                        bg="powder blue", fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)
        
        #=====Variables=====
        shid = StringVar()
        indate = StringVar(value=datetime.date.today().strftime('%Y-%m-%d'))
        saledate = StringVar(value=datetime.date.today().strftime('%Y-%m-%d'))
        stype = StringVar()
        customer = StringVar()
        sterm = StringVar()
        coa = StringVar()
        total = StringVar()
        vat = StringVar()
        net = StringVar()
        currency = StringVar()
        exrate = StringVar()
        gubun = StringVar()
        agroup = StringVar()
        descript = StringVar()
        
    #=======upperframe Functions=========================
        #Variable declare and the maximum value of the sdid field from the sales_detail table
        mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
        cursor=mydb.cursor()
        cursor.execute("SELECT MAX(shid) FROM sales_header")
        result_1 = cursor.fetchone()
        if result_1 is not None:
            max_shid = result_1[0]
        else:
            max_shid = 0  # Assign a default value of 0 if max_tsq is None
        shid_var = StringVar()
        shid_var.set(str(max_shid + 1))
        
        def fetch_header():
            try:
                # MySQL 데이터베이스 연결
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                cursor.execute("SELECT * FROM sales_header")
                result = cursor.fetchall()
                return result
            except pymysql.Error as e:
                print("MySQL Error:", e)
            finally: # 연결 종료
                cursor.close()
                mydb.close()
        
        def add_to_treeview():
            rows = fetch_header()
            if rows is not None:
                tree.delete(*tree.get_children())
                for row in rows:
                    tree.insert('', tk.END, values=row)

        def clear_header(*clicked):
            if clicked:
                tree.selection_remove(tree.focus())
                tree.focus('')
                
            txtshid.delete(0, END)
            txtindate.delete(0, END)
            txtsaledate.delete(0, END)
            combstype.set('')
            combcustomer.set('')
            combsterm.set('')
            combcoa.set('')
            txttotal.delete(0, END)
            txtvat.delete(0, END)
            txtnet.delete(0, END)
            txtcurrency.delete(0, END)
            txtexrate.delete(0, END)
            combgubun.set('')
            combagroup.set('')
            txtdescript.delete(0, END)
        
        def display_header(event):
            selected_item = tree.focus()
            if selected_item:
                row = tree.item(selected_item)['values']
                clear_header()
                txtshid.delete(0, 'end')
                txtshid.insert(0, row[0])
                txtindate.delete(0, 'end')
                txtindate.insert(0, row[1])
                txtsaledate.delete(0, 'end')
                txtsaledate.insert(0, row[2])
                combstype.set(row[3])
                combcustomer.set(row[4])
                combsterm.set(row[5])
                combcoa.set(row[6])
                txttotal.delete(0, 'end')
                txttotal.insert(0, row[7])
                txtvat.delete(0, 'end')
                txtvat.insert(0,row[8])
                txtnet.delete(0, 'end')
                txtnet.insert(0,str(row[9]))  # Adjusted index to match the column
                txtcurrency.delete(0, 'end')
                txtcurrency.insert(0, row[10])  # Adjusted index to match the column
                txtexrate.delete(0, 'end')
                txtexrate.insert(0, str(row[11]))  # Convert decimal to string with two decimal places
                combgubun.set(row[12])  # Adjusted index to match the column
                combagroup.set(row[13])  # Adjusted index to match the column
                txtdescript.delete(0, 'end')
                txtdescript.insert(0, row[14])  # Adjusted index to match the column
                
            else:
                pass
        
        def add_header():
            # Get values from entry widgets
            shid = txtshid.get()
            indate = txtindate.get()
            saledate = txtsaledate.get()
            stype = combstype.get()
            customer = combcustomer.get()
            sterm = combsterm.get()
            coa = combcoa.get()
            total = txttotal.get()
            vat = txtvat.get()
            net = txtnet.get()
            currency = txtcurrency.get()
            exrate = txtexrate.get()
            gubun = combgubun.get()
            agroup = combagroup.get()
            descript = txtdescript.get()

            try:
                # Connect to the database
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()

                # Insert data into the sales_header table
                query = "INSERT INTO sales_header ( indate, saledate, stype, customer, sterm, coa, total, vat, net, currency, exrate, gubun, agroup, descript)" \
                                            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s)"
                val = (indate, saledate, stype, customer, sterm, coa, total, vat, net, currency, exrate, gubun, agroup, descript)
                cursor.execute(query, val) #(indate, saledate, stype, customer, sterm, coa, total, vat, net, currency, exrate, gubun, agroup, descript))

                # Commit changes and close connection
                mydb.commit()
                mydb.close()
                messagebox.showinfo('Success', 'Data has been inserted.')
                add_to_treeview()
                # Clear entry fields after successful insertion
                clear_header()

            except Exception as e:
                print("Error:", e)
                mydb.rollback()
                mydb.close()
                messagebox.showerror("Error", "Data not inserted successfully...") 

        def update_header():
            selected_item = tree.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to update.')
            else:
                # Get values from entry widgets
                shid = txtshid.get()
                indate = txtindate.get()
                saledate = txtsaledate.get()
                stype = combstype.get()
                customer = combcustomer.get()
                sterm = combsterm.get()
                coa = combcoa.get()
                total = txttotal.get()
                vat = txtvat.get()
                net = txtnet.get()
                currency = txtcurrency.get()
                exrate = txtexrate.get()
                gubun = combgubun.get()
                agroup = combagroup.get()
                descript = txtdescript.get()

                try:
                    # Connect to the database
                    mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                    cursor = mydb.cursor()

                    # Update data in the sales_header table
                    query = "UPDATE sales_header SET indate=%s, saledate=%s, stype=%s, customer=%s, sterm=%s, coa=%s, total=%s, vat=%s, net=%s, currency=%s," \
                            "exrate=%s, gubun=%s, agroup=%s, descript=%s WHERE shid=%s"
                    cursor.execute(query, (indate, saledate, stype, customer, sterm, coa, total, vat, net, currency, exrate, gubun, agroup, descript, shid))

                    # Commit changes and close connection
                    mydb.commit()
                    mydb.close()
                    lastid = cursor.lastrowid
                    add_to_treeview()
                    messagebox.showinfo('Success', 'Data has been updated.')
                    # Clear entry fields after successful update
                    clear_header()

                except Exception as e:
                    print("Error:", e)
        
        def delete_header():
            selected_item = tree.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to delete.')
            else:
                shid = txtshid.get()
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "DELETE from sales_header where shid=%s"
                    val = (shid,) #'val' 변수가 튜플이 아닌 함수로 설정되어 있기 때문입니다. 'val,' 변수를 튜플로 설정해야 합니다.
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    add_to_treeview()
                    clear_header()
                    messagebox.showinfo('Success', 'Data has been deleted.')

                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not deleted successfully...")
        
        #========MS EXCEL INVOICE FORM MAKING===============
    
        # 인보이스 폼 생성 함수
        def create_invoice(shid_value):
            # 새로운 엑셀 워크북 생성
            workbook = Workbook()
            sheet = workbook.active

            # sales_header 테이블에서 해당 shid의 정보 가져오기
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            query = "SELECT * FROM sales_header WHERE shid = %s"
            cursor.execute(query, (shid_value,))
            header = cursor.fetchone()

            if header:
                # 인보이스 공통 정보 엑셀에 입력
                sheet.cell(row=1, column=1, value="Invoice ID") 
                sheet.cell(row=1, column=2, value=header[0])
                sheet.cell(row=2, column=1, value="Customer Name")
                sheet.cell(row=2, column=2, value=header[1])
                sheet.cell(row=3, column=1, value="Invoice Date")
                sheet.cell(row=3, column=2, value=header[2])
                sheet.cell(row=4, column=1, value="Total Amount")
                sheet.cell(row=4, column=2, value=header[3])

                # sales_detail 테이블에서 해당 shid의 아이템 정보 가져오기
                query = "SELECT * FROM sales_detail WHERE shid = %s"
                cursor.execute(query, (shid_value,))
                details = cursor.fetchall()

                if details:
                    # 인보이스 아이템 정보 엑셀에 입력
                    sheet.cell(row=6, column=1, value="SQ")
                    sheet.cell(row=6, column=2, value="Product")
                    sheet.cell(row=6, column=3, value="Item")  
                    sheet.cell(row=6, column=4, value="Quantity")
                    sheet.cell(row=6, column=5, value="Kg")
                    sheet.cell(row=6, column=6, value="@Price")
                    sheet.cell(row=6, column=7, value="Amount")
                    sheet.cell(row=6, column=8, value="Currency")
                    sheet.cell(row=6, column=9, value="Exrate")
                    sheet.cell(row=6, column=10, value="Memo")

                    row_num = 7
                    for detail in details:
                        sheet.cell(row=row_num, column=1, value=detail[1])
                        sheet.cell(row=row_num, column=2, value=detail[2])
                        sheet.cell(row=row_num, column=3, value=detail[3])
                        sheet.cell(row=row_num, column=4, value=detail[4])
                        sheet.cell(row=row_num, column=5, value=detail[5])
                        sheet.cell(row=row_num, column=6, value=detail[6])
                        sheet.cell(row=row_num, column=7, value=detail[7])
                        sheet.cell(row=row_num, column=8, value=detail[8])
                        sheet.cell(row=row_num, column=9, value=detail[9])
                        sheet.cell(row=row_num, column=10, value=detail[10])
                        row_num += 1
                else:
                    sheet.cell(row=6, column=1, value="No invoice items found.")

                today = datetime.today().strftime('%Y%m%d')
                
                # 엑셀 파일 저장
                workbook.save(f"D:\\15Work\\fkldb\\2023\\{today}_invoice_{shid_value}.xlsx")
                filename = f"D:\\15Work\\fkldb\\2023\\{today}_invoice_{shid_value}.xlsx"
                print(f"Invoice_{shid_value} created successfully.")
                #os.system('start excel.exe invoice.xlsx')
                os.startfile(filename)
            else:
                print("Invoice not found.")

            mydb.close()
                    
        def run_invoice_form():
            shid_value = txtshid.get()  # txtshid에서 값을 가져옵니다.
            # # subprocess를 사용하여 tk_invoiceform.py 스크립트를 실행하고, shid 값을 인자로 전달합니다.
            # subprocess.run(["python", "D:\\20Program\\fkl_acc\\tk_invoiceform.py", shid_value], check=True)
            create_invoice(shid_value)       
        
        #===========upper Frame=======================
        upperframe = Frame(parent, bd=3, relief=RIDGE,padx=20, bg="powder blue")
        upperframe.place(x=0, y=50, width=1700, height=423)
        
        #=====Sales_Header_Input====
        salesheader=LabelFrame(upperframe, text="Sales Invoice Header Input", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=10)
        salesheader.place(x=0, y=2, width=760, height=280)
        
        lblshid = Label(salesheader, text="Sales ID:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblindate = Label(salesheader, text="Invoice Date:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblsaledate = Label(salesheader, text="Sales Date:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblstype = Label(salesheader, text="Sale Type:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblcustomer = Label(salesheader, text="Customer:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblsterm = Label(salesheader, text="Sales Term:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblcoa = Label(salesheader, text="COA:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lbltotal = Label(salesheader, text="Total:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblvat = Label(salesheader, text="Vat:", font=("Arial", 10, "bold"), padx=1, pady=3, bg="powder blue")
        lblnet = Label(salesheader, text="Net:", font=("Arial", 10, "bold"), padx=5, pady=3, bg="powder blue")
        lblcurrency = Label(salesheader, text="Currency:", font=("Arial", 10, "bold"), padx=5, pady=3, bg="powder blue")
        lblexrate = Label(salesheader, text="Exchange Rate:", font=("Arial", 10, "bold"), padx=5, pady=3, bg="powder blue")
        lblgubun = Label(salesheader, text="Gubun:", font=("Arial", 10, "bold"), padx=3, pady=5, bg="powder blue")
        lblagroup = Label(salesheader, text="Account Group:", font=("Arial", 10, "bold"), padx=5, pady=3, bg="powder blue")
        lbldescript = Label(salesheader, text="Description:", font=("Arial", 10, "bold"), padx=5, pady=3, bg="powder blue")
        
        lblshid.grid(row=0, column=0, sticky=W)
        lblindate.grid(row=1, column=0, sticky=W)
        lblsaledate.grid(row=2, column=0, sticky=W)
        lblstype.grid(row=3, column=0, sticky=W)
        lblcustomer.grid(row=4, column=0, sticky=W)
        lblsterm.grid(row=5, column=0, sticky=W)
        lblcoa.grid(row=6, column=0, sticky=W)
        lbltotal.grid(row=7, column=0, sticky=W)
        lblvat.grid(row=8, column=0, sticky=W)
        lblnet.grid(row=0, column=2, sticky=W)
        lblcurrency.grid(row=1, column=2, sticky=W)
        lblexrate.grid(row=2, column=2, sticky=W)
        lblgubun.grid(row=3, column=2, sticky=W)
        lblagroup.grid(row=4, column=2, sticky=W)
        lbldescript.grid(row=5, column=2, sticky=W)
        
        txtshid = Entry(salesheader, font=("Arial", 12, "bold"), textvariable=shid_var, width=29)
        txtindate = Entry(salesheader, font=("Arial", 12, "bold"), textvariable=indate, width=29)
        txtsaledate = Entry(salesheader, font=("Arial", 12, "bold"), textvariable=saledate, width=29)
        combstype = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=stype, width=27, state="readonly")
        combstype['value'] = ("Domestic", "Overseas", "Others")
        combstype.current(0)
        combcustomer = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=customer, width=27, state="readonly")
        combcustomer['values'] = self.company_names()
        combcustomer.current(0)  # Optionally set the default selection
        #combcustomer.pack()
        combsterm = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=sterm, width=27, state="readonly")
        combsterm['value'] = ("EOD", "FOB", "CIF", "EXW", "DDR")
        combsterm.current(0)
        combcoa = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=coa, width=27, state="readonly")
        combcoa['values'] = self.coa_accounts()
        combcoa.current(0)
        txttotal = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        txtvat = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        txtnet = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        txtcurrency = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        txtexrate = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        combgubun = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=gubun, width=27, state="readonly")
        combgubun['value'] = ('CASH', 'BANK', 'CreditCard', 'A/P', 'ACCRUD','ETC')
        combagroup = ttk.Combobox(salesheader, font=("Arial", 12, "bold"), textvariable=agroup, width=27, state="readonly")
        combagroup['value'] = ('자산구매', '자산처분', '감가상각', '부채발생', '부채감소', '대표대여금', '세금계산서', '상품구매','기타원가항목','기타')
        txtdescript = Entry(salesheader, font=("Arial", 12, "bold"), width=29)
        
        txtshid.grid(row=0, column=1)
        txtindate.grid(row=1, column=1)
        txtsaledate.grid(row=2, column=1)
        combstype.grid(row=3, column=1)
        combcustomer.grid(row=4, column=1)
        combsterm.grid(row=5, column=1)
        combcoa.grid(row=6, column=1)
        txttotal.grid(row=7, column=1)
        txtvat.grid(row=8, column=1)
        txtnet.grid(row=0, column=3)
        txtcurrency.grid(row=1, column=3)
        txtexrate.grid(row=2, column=3)
        combgubun.grid(row=3, column=3)
        combagroup.grid(row=4, column=3)
        txtdescript.grid(row=5, column=3)
        
        #=====ButtonFrame=====
        buttonframe1 = LabelFrame(upperframe, bd=3, relief=RIDGE, padx=2, bg="powder blue")
        buttonframe1.place(x=0, y=285, width=560 , height=40)
        
        btnAddData = Button(buttonframe1, text="Add Data", command=add_header, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData.grid(row=0, column=0)
        
        btnAddData = Button(buttonframe1, text="Update", command=update_header, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData.grid(row=0, column=1)
        
        btnAddData = Button(buttonframe1, text="Delete", command=delete_header, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData.grid(row=0, column=2)
        
        btnAddData = Button(buttonframe1, text="Reset", command="", font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData.grid(row=0, column=3)
        
        btnAddData = Button(buttonframe1, text="Invoice", command=run_invoice_form, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData.grid(row=0, column=4)
        
        #=====Sales_Header_DB=====
        salesheader_db = LabelFrame(upperframe, text="Sales Invoice Header DB", bg="powder blue", bd=2, relief=RIDGE, font=("Arial", 12, "bold"), padx=2)
        salesheader_db.place(x=765, y=2, width=900, height=405)

        tableframe=Frame(salesheader_db, bd=2, relief=RIDGE, bg="powder blue")
        tableframe.place(x=0,y=2,width=895, height=380)
        
        xscroll=Scrollbar(tableframe, orient=HORIZONTAL)
        Yscroll=Scrollbar(tableframe, orient=VERTICAL)
        
        tree=ttk.Treeview(tableframe, columns=("shid", "indate", "saledate", "stype", "customer", "sterm", "coa", "total", "vat", "net",
                                            "currency","exrate","gubun","agroup", "descript"), xscrollcommand=xscroll.set, yscrollcommand=Yscroll.set)
        
        xscroll.pack(side=BOTTOM, fill=X)
        Yscroll.pack(side=RIGHT, fill=Y)
        xscroll.config(command=tree.xview)
        Yscroll.config(command=tree.yview)
        
        tree.heading("shid", text="Sale sq")
        tree.heading("indate", text="InputDate")
        tree.heading("saledate", text="SalesDate")
        tree.heading("stype", text="SaleType")
        tree.heading("customer", text="Customer")
        tree.heading("sterm", text="SalesTerm")
        tree.heading("coa", text="COA")
        tree.heading("total", text="Total")
        tree.heading("vat", text="VAT")
        tree.heading("net", text="Net")
        tree.heading("currency", text="Currency")
        tree.heading("exrate", text="ExRate")
        tree.heading("gubun", text="Gubun")
        tree.heading("agroup", text="AGroup")
        tree.heading("descript", text="Description")
        
        tree["show"]="headings"
        tree.pack(fill=BOTH, expand=1)
        
        # tree.column("sid", width=100)
        
        tree.bind("<ButtonRelease>", display_header)
        add_to_treeview()
        
    #===========lowerFrame Function=================================
        #Variable declare and the maximum value of the sdid field from the sales_detail table
        mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
        cursor=mydb.cursor()
        cursor.execute("SELECT MAX(sdid) FROM sales_detail")
        result_2 = cursor.fetchone()
        if result_2 is not None:
            max_sdid = result_2[0]
        else:
            max_sdid = 0  # Assign a default value of 0 if max_tsq is None
        sdid_var = StringVar()
        sdid_var.set(str(max_sdid + 1))
                
        
        def fetch_detail():
            try:
                # MySQL 데이터베이스 연결
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()
                cursor.execute("SELECT * FROM sales_detail")
                result = cursor.fetchall()
                return result
            except pymysql.Error as e:
                print("MySQL Error:", e)
            finally: # 연결 종료
                cursor.close()
                mydb.close()
        
        def add_to_treeview_2():
            rows = fetch_detail()
            if rows is not None:
                tree_2.delete(*tree_2.get_children())
                for row in rows:
                    tree_2.insert('', tk.END, values=row)

        def clear_detail(*clicked):
            if clicked:
                tree_2.selection_remove(tree_2.focus())
                tree_2.focus('')

            txtsdid.delete(0, END)
            txtshid_2.delete(0, END)
            txtproduct.delete(0, END)
            txtitem.delete(0, END)
            txtqty.delete(0, END)
            txtkg.delete(0, END)
            txtunitprice.delete(0, END)
            txtnet_2.delete(0, END)
            txtcurrency_2.delete(0, END)
            txtexrate_2.delete(0, END)
            txtmemo.delete(0, END)
        
        def display_detail(event):
            selected_item = tree_2.focus()
            if selected_item:
                row = tree_2.item(selected_item)['values']
                clear_detail()
                txtsdid.delete(0, 'end')
                txtsdid.insert(0, row[0])
                txtshid_2.delete(0, 'end')
                txtshid_2.insert(0, row[1])
                txtproduct.delete(0, 'end')
                txtproduct.insert(0, row[2])
                txtitem.delete(0, 'end')
                txtitem.insert(0, row[3])
                txtqty.delete(0, 'end')
                txtqty.insert(0, row[4])
                txtkg.delete(0, 'end')
                txtkg.insert(0, row[5])
                txtunitprice.delete(0, 'end')
                txtunitprice.insert(0, row[6])
                txtnet_2.delete(0, 'end')
                txtnet_2.insert(0, row[7])
                txtcurrency_2.delete(0, 'end')
                txtcurrency_2.insert(0, row[8])
                txtexrate_2.delete(0, 'end')
                txtexrate_2.insert(0, row[9])
                txtmemo.delete(0, 'end')
                txtmemo.insert(0, row[10])
                
            else:
                pass
        
        def add_detail():
            # Get values from entry widgets
            #sdid = txtsdid.get()
            shid = txtshid_2.get()
            product = txtproduct.get()
            item = txtitem.get()
            qty = txtqty.get()
            kg = txtkg.get()
            unitprice = txtunitprice.get()
            net_2 = txtnet_2.get()
            currency_2 = txtcurrency_2.get()
            exrate_2 = txtexrate_2.get()
            memo = txtmemo.get()

            try:
                # Connect to the database
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                cursor = mydb.cursor()

                # Insert data into the sales_header table
                query = "INSERT INTO sales_detail (shid, product, item, qty, kg, unitprice, net, currency, exrate, memo)" \
                                            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                val = (shid, product, item, qty, kg, unitprice, net_2, currency_2, exrate_2, memo)
                cursor.execute(query, val)

                # Commit changes and close connection
                mydb.commit()
                mydb.close()
                messagebox.showinfo('Success', 'Data has been inserted.')
                add_to_treeview_2()
                # Clear entry fields after successful insertion
                clear_detail()

            except Exception as e:
                print("Error:", e)
                mydb.rollback()
                mydb.close()
                messagebox.showerror("Error", "Data not inserted successfully...") 


        def update_detail():
            selected_item = tree_2.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to update.')
            else:
                # Get values from entry widgets
                sdid = txtsdid.get()
                shid = txtshid_2.get()
                product = txtproduct.get()
                item = txtitem.get()
                qty = txtqty.get()
                kg = txtkg.get()
                unitprice = txtunitprice.get()
                net_2 = txtnet_2.get()
                currency_2 = txtcurrency_2.get()
                exrate_2 = txtexrate_2.get()
                memo = txtmemo.get()

                try:
                    # Connect to the database
                    mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
                    cursor = mydb.cursor()

                    # Update data in the sales_header table
                    query = "UPDATE sales_detail SET shid=%s, product=%s, item=%s, qty=%s, kg=%s, unitprice=%s, net=%s, currency=%s, exrate=%s, memo=%s" \
                            "WHERE sdid=%s"
                    val = (shid, product, item, qty, kg, unitprice, net_2, currency_2, exrate_2, memo, sdid)
                    cursor.execute(query, val)

                    # Commit changes and close connection
                    mydb.commit()
                    mydb.close()
                    lastid = cursor.lastrowid
                    add_to_treeview_2()
                    messagebox.showinfo('Success', 'Data has been updated.')
                    # Clear entry fields after successful update
                    #clear_detail()

                except Exception as e:
                    print("Error:", e)
                        
        
        def delete_detail():
            selected_item = tree_2.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to delete.')
            else:
                sdid = txtsdid.get()
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "DELETE from sales_detail where sdid=%s"
                    val = (sdid,) #'val' 변수가 튜플이 아닌 함수로 설정되어 있기 때문입니다. 'val,' 변수를 튜플로 설정해야 합니다.
                    cursor.execute(sql, val)
                    mydb.commit()
                    lastid = cursor.lastrowid
                    add_to_treeview_2()
                    clear_detail()
                    messagebox.showinfo('Success', 'Data has been deleted.')

                except Exception as e:
                    print(e)
                    mydb.rollback()
                    mydb.close()
                    messagebox.showerror("Error", "Data not deleted successfully...")
                    
        def calculate_net_sum(txtshid_value):
            try:
                # Connect to the database
                mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl",charset='utf8mb4')
                cursor = mydb.cursor()
                # Execute the SQL query to calculate the sum of 'net' values
                query = "SELECT SUM(net) FROM sales_detail WHERE shid = %s"
                cursor.execute(query, (txtshid_value,))
                
                # # 결과 가져오기
                # result = cursor.fetchone()
                # if result:
                #     net_sum = float(result[0])
                # else:
                #     net_sum = 0
                
                net_sum = cursor.fetchone()  # Fetch the sum of 'net' values
                net_sum = str(net_sum).replace('','')  # Convert the string to float and remove the comma
                # If there are no records, set the sum to 0
                if net_sum is None:
                    net_sum = 0
                return float(net_sum)
            
            except Exception as e:
                print("Error:", e)
            finally:
                cursor.close()
                mydb.close()

        def on_add_data_click():
            # Get the value of txtshid
            txtshid_value = txtshid_2.get()
            # Calculate the sum of 'net' values
            total_net = calculate_net_sum(txtshid_value)
            # Update the value of txtnet
            txtnet.delete(0, END)
            txtnet.insert(0, str(total_net))
        
        def calculate_net():
            qty = txtqty.get()
            kg = txtkg.get()
            unit_price = txtunitprice.get()

            if qty != '0' and qty != '':
                net = float(qty) * float(unit_price)
                txtnet_2.delete(0, END)
                txtnet_2.insert(0, str(net))
            elif kg != '0' and kg != '':
                net = float(kg) * float(unit_price)
                txtnet_2.delete(0, END)
                txtnet_2.insert(0, str(net))        
        
    #=======lowerframe=================================================================    
        lowerframe = Frame(parent, bd=2, relief=RIDGE,padx=20, bg="powder blue")
        lowerframe.place(x=0, y=480, width=1700, height=423)
        
        #=====Sales Detail Input=====
        salesdetail = LabelFrame(lowerframe, text="Sales Invoice Detail Input", bg="powder blue", bd=2, relief=RIDGE, font=("Arial", 12, "bold"), padx=10)
        salesdetail.place(x=0, y=2, width=760, height=280)

        lblsdid = Label(salesdetail, text="Detail ID:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblshid_2 = Label(salesdetail, text="Sales ID:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblproduct = Label(salesdetail, text="Products:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblitem = Label(salesdetail, text="Item:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblqty = Label(salesdetail, text="Qty:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblkg = Label(salesdetail, text="Kg:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblunitprice = Label(salesdetail, text="UnitPrice:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblnet_2 = Label(salesdetail, text="Net:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblcurrency_2 = Label(salesdetail, text="Currency:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblexrate_2 = Label(salesdetail, text="ExRate:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        lblmemo = Label(salesdetail, text="Memo:", font=("Arial", 10, "bold"), padx=2, pady=6, bg="powder blue")
        
        lblsdid.grid(row=0, column=0, sticky=W)
        lblshid_2.grid(row=1, column=0, sticky=W)
        lblproduct.grid(row=2, column=0, sticky=W)
        lblitem.grid(row=3, column=0, sticky=W)
        lblqty.grid(row=4, column=0, sticky=W)
        lblkg.grid(row=5, column=0, sticky=W)
        lblunitprice.grid(row=0, column=2, sticky=W)
        lblnet_2.grid(row=1, column=2, sticky=W)
        lblcurrency_2.grid(row=2, column=2, sticky=W)
        lblexrate_2.grid(row=3, column=2, sticky=W)
        lblmemo.grid(row=4, column=2, sticky=W)
        
        txtsdid = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtshid_2 = Entry(salesdetail, textvariable=shid_var ,font=("Arial", 12, "bold"), width=29)
        txtproduct = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtitem = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtqty = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtkg = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtunitprice = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtunitprice.bind('<KeyRelease>', lambda event: calculate_net())
        txtnet_2 = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtcurrency_2 = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtexrate_2 = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        txtmemo = Entry(salesdetail, font=("Arial", 12, "bold"), width=29)
        
        txtsdid.grid(row=0, column=1)
        txtshid_2.grid(row=1, column=1)
        txtproduct.grid(row=2, column=1)
        txtitem.grid(row=3, column=1)
        txtqty.grid(row=4, column=1)
        txtkg.grid(row=5, column=1)
        txtunitprice.grid(row=0, column=3)
        txtnet_2.grid(row=1, column=3)
        txtcurrency_2.grid(row=2, column=3)
        txtexrate_2.grid(row=3, column=3)
        txtmemo.grid(row=4, column=3)
        
        #=====ButtonFrame=====
        buttonframe2 = LabelFrame(lowerframe, bd=3, relief=RIDGE, padx=2, bg="powder blue")
        buttonframe2.place(x=0, y=285, width=560 , height=40)
        
        btnAddData_2 = Button(buttonframe2, text="Add Data", command=add_detail, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData_2.grid(row=0, column=0)
        
        btnAddData_2 = Button(buttonframe2, text="Update", command=update_detail, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData_2.grid(row=0, column=1)
        
        btnAddData_2 = Button(buttonframe2, text="Delete", command=delete_detail, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData_2.grid(row=0, column=2)
        
        btnAddData_2 = Button(buttonframe2, text="Reset", command=clear_detail, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData_2.grid(row=0, column=3)
        
        btnAddData_2 = Button(buttonframe2, text="Sum(Net)", command=on_add_data_click, font=("Arial", 12, "bold"), width=10, bg="blue", fg="white")
        btnAddData_2.grid(row=0, column=4)
        
        #=====Sales_Header_DB=====
        salesdetail_db = LabelFrame(lowerframe, text="Sales Invoice Detail DB", bg="powder blue", bd=2, relief=RIDGE, font=("Arial", 12, "bold"), padx=2)
        salesdetail_db.place(x=765, y=2, width=900, height=405)

        tableframe_2=Frame(salesdetail_db, bd=2, relief=RIDGE, bg="powder blue")
        tableframe_2.place(x=0,y=2,width=895, height=380)
        
        xscroll=Scrollbar(tableframe_2, orient=HORIZONTAL)
        Yscroll=Scrollbar(tableframe_2, orient=VERTICAL)
        
        tree_2=ttk.Treeview(tableframe_2, columns=("sdid", "shid", "product", "item", "qty", "kg", "unitprice", "net", "currency", 
                                                            "exrate","memo"), xscrollcommand=xscroll.set, yscrollcommand=Yscroll.set)
        
        xscroll.pack(side=BOTTOM, fill=X)
        Yscroll.pack(side=RIGHT, fill=Y)
        xscroll.config(command=tree_2.xview)
        Yscroll.config(command=tree_2.yview)
        
        tree_2.heading("sdid", text="SDID")
        tree_2.heading("shid", text="SHID")
        tree_2.heading("product", text="Product")
        tree_2.heading("item", text="Item")
        tree_2.heading("qty", text="Qty")
        tree_2.heading("kg", text="Kg")
        tree_2.heading("unitprice", text="@Price")
        tree_2.heading("net", text="Net")
        tree_2.heading("currency", text="Currency")
        tree_2.heading("exrate", text="ExRate")
        tree_2.heading("memo", text="Memo")
        
        tree_2["show"]="headings"
        tree_2.pack(fill=BOTH, expand=1)
        
        tree_2.bind("<ButtonRelease>", display_detail)        
        add_to_treeview_2()
    
    #+++++++++++Global Functions++++++++++++++++++++++
    # Function to get today's date in the format "yyyymmdd"
    def get_today_date(self):
        return datetime.today().strftime('%Y-%m-%d')

    # Connect to the database and fetch the company names
    def company_names(self):
        mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = mydb.cursor()
        cursor.execute("SELECT company FROM customer")
        companies = cursor.fetchall()
        # Extract company names from the fetched data
        company_names = [company[0] for company in companies]
        return company_names  # Return the list of company names
    
    def coa_accounts(self):
        mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = mydb.cursor()
        cursor.execute("SELECT COA2 FROM coa")
        coas = cursor.fetchall()
        # Extract company names from the fetched data
        coa_accounts = [coa[0] for coa in coas]
        return coa_accounts  # Return the list of company names
    
#========== 여기부터 코드 작성 ==========#
class Purchase(tk.Frame):  
    def __init__(self, parent):
        super().__init__(parent, bg='powder blue')
        self.pack(fill=tk.BOTH, expand=True)
        lbltitle = tk.Label(self, text="Purchase Invoice Management System", bd=3, relief=RIDGE, bg="powder blue", 
                        fg="black", font=("Arial", 14, "bold"), padx=2, pady=10)
        lbltitle.pack(side=TOP, fill=X)
        
        pid = StringVar()
        linesq = StringVar()
        supplier = StringVar()
        indate = StringVar()
        voudate = StringVar()
        product = StringVar()
        item = StringVar()
        qty = StringVar()
        kg = StringVar()
        coa = StringVar()
        amount = StringVar()
        vat = StringVar()
        exrate = StringVar()
        currency = StringVar()
        descript = StringVar()
        agroup = StringVar()
        gubun = StringVar()
        vouno = StringVar()
        memo = StringVar()


        opt_gubun = ['자산구매', '자산처분', '감가상각', '부채발생', '부채감소', '대표대여금', '세금계산서', '상품구매','기타원가항목','기타']
        Var_gubun = StringVar()
        opt_agroup = ['CASH', 'BANK', 'CreditCard', 'A/P', 'ACCRUD','ETC']
        Var_agroup = StringVar()

        opt_productgroup = ["Cup Pipe", "CuBusbar","ALBusbar","ALProfile", "Machines", "Tools"]
        Var_productgroup = StringVar()

        #=====Supplier Frame=====
        supplier_frame=LabelFrame(parent, text="Supplier Group Selection", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=2, pady=1)
        supplier_frame.place(x=0, y=50, width=570, height=450)

        comboproductgroup = ttk.Combobox(supplier_frame, textvariable=Var_productgroup, values=opt_productgroup, font=("Arial", 12, "bold"), state="readonly")
        comboproductgroup['values'] = ("Cup Pipe", "CuBusbar","ALBusbar","ALProfile", "Machines", "Tools")
        comboproductgroup.grid(row=1, column=0, padx=10, pady=10, sticky="w")


        mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
        cursor = mydb.cursor()

        def update_supplier_list(event):
            selected_product = comboproductgroup.get()
            if selected_product:
                cursor.execute("SELECT pid, supplier FROM cogs WHERE product=%s", (selected_product,))
                rows = cursor.fetchall()
                supplier_listbox.delete(0, END)
                for row in rows:
                    supplier_listbox.insert(END, f"{row[0]} - {row[1]}")
            mydb.commit()

        comboproductgroup.bind("<<ComboboxSelected>>", update_supplier_list)

        supplier_listbox = Listbox(supplier_frame, font=("Arial", 12), width=59, height=19)
        supplier_listbox.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        def select_supplier(event):
            selected_index = supplier_listbox.curselection()
            if selected_index:
                selected_row = supplier_listbox.get(selected_index[0])
                pid, supplier_name = selected_row.split(" - ")
                txtsupplier.delete(0, END)
                txtsupplier.insert(END, supplier_name)
                comboproduct.set(pid)

        supplier_listbox.bind("<Double-Button-1>", select_supplier)

        #=====Entries Frame=====
        entries_frame=LabelFrame(parent, text="Purchasing Input Management System", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=20)
        entries_frame.place(x=600, y=50, width=870, height=450)

        lblpid = Label(entries_frame, text="ID", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lbllinesq = Label(entries_frame, text="LineSq", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblsupplier = Label(entries_frame, text="Supplier", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblindate = Label(entries_frame, text="Indate", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblvoudate = Label(entries_frame, text="VouDate", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblproduct = Label(entries_frame, text="Product", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblitem = Label(entries_frame, text="Item", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblqty = Label(entries_frame, text="Qty", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblkg = Label(entries_frame, text="Kg", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcoa = Label(entries_frame, text="Coa", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblamount = Label(entries_frame, text="Amount", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblvat = Label(entries_frame, text="Vat", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblexrate = Label(entries_frame, text="ExRate", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblcurrency = Label(entries_frame, text="Currency", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lbldescript = Label(entries_frame, text="Descript", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblagroup = Label(entries_frame, text="AGroup", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblgubun = Label(entries_frame, text="Gubun", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblvouno = Label(entries_frame, text="VouNo", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")
        lblmemo = Label(entries_frame, text="Memo", font=("Arial", 12, "bold"), padx=2, pady=6, bg="powder blue", fg="green")

        lblpid.grid(row=0, column=0, sticky="w")
        lbllinesq.grid(row=1, column=0, sticky="w")
        lblsupplier.grid(row=2, column=0, sticky="w")
        lblindate.grid(row=3, column=0, sticky="w")
        lblvoudate.grid(row=4, column=0, sticky="w")
        lblproduct.grid(row=5, column=0, sticky="w")
        lblitem.grid(row=6, column=0, sticky="w")
        lblqty.grid(row=7, column=0, sticky="w")
        lblkg.grid(row=8, column=0, sticky="w")
        lblcoa.grid(row=9, column=0, sticky="w")
        lblamount.grid(row=0, column=2, padx=30,sticky=W)
        lblvat.grid(row=1, column=2, padx=30,sticky=W)
        lblexrate.grid(row=2, column=2, padx=30,sticky=W)
        lblcurrency.grid(row=3, column=2, padx=30,sticky=W)
        lbldescript.grid(row=4, column=2, padx=30,sticky=W)
        lblagroup.grid(row=5, column=2, padx=30,sticky=W)
        lblgubun.grid(row=6, column=2, padx=30, sticky=W)
        lblvouno.grid(row=7, column=2, padx=30,sticky=W)
        lblmemo.grid(row=8, column=2, padx=30,sticky=W)

        txtpid = Entry(entries_frame, textvariable=pid, font=("Arial", 12, "bold"), width=30)
        txtlinesq = Entry(entries_frame, textvariable=linesq, font=("Arial", 12, "bold"), width=30)
        txtsupplier = Entry(entries_frame, textvariable=supplier, font=("Arial", 12, "bold"), width=30)
        txtindate = Entry(entries_frame, textvariable=indate, font=("Arial", 12, "bold"), width=30)
        txtvoudate = Entry(entries_frame, textvariable=indate, font=("Arial", 12, "bold"), width=30)
        comboproduct = ttk.Combobox(entries_frame, textvariable=product, font=("Arial", 12, "bold"), width=28, state="readonly")
        comboproduct['values'] = ("Cup Pipe", "CuBusbar","ALBusbar","ALProfile", "Machines", "Tools", "ETC")
        txtitem = Entry(entries_frame, textvariable=item, font=("Arial", 12, "bold"), width=30)
        txtqty = Entry(entries_frame, textvariable=qty, font=("Arial", 12, "bold"), width=30)
        txtkg = Entry(entries_frame, textvariable=kg, font=("Arial", 12, "bold"), width=30)

        cursor.execute("select coa1 from coa")
        coa_data = cursor.fetchall()
        coa_values = [row[0] for row in coa_data]
        combocoa = ttk.Combobox(entries_frame,textvariable=coa,font=("Arial", 12, "bold"), width=28 ,state='readonly')
        combocoa['values'] = coa_values
        combocoa.current(53)

        txtamount = Entry(entries_frame, textvariable=amount, font=("Arial", 12, "bold"), width=30)
        txtvat = Entry(entries_frame, textvariable=vat, font=("Arial", 12, "bold"), width=30)
        txtexrate = Entry(entries_frame, textvariable=exrate, font=("Arial", 12, "bold"), width=30)
        txtcurrency = Entry(entries_frame, textvariable=currency, font=("Arial", 12, "bold"), width=30)
        txtdescript = Entry(entries_frame, textvariable=descript, font=("Arial", 12, "bold"), width=30)
        comboagroup = ttk.Combobox(entries_frame, textvariable=Var_agroup, values=opt_agroup, font=("Arial", 12, "bold"), width=28, state="readonly")
        comboagroup['values'] = ('CASH', 'BANK', 'CreditCard', 'A/P', 'ACCRUD','ETC')
        combogubun = ttk.Combobox(entries_frame, textvariable=Var_gubun, values=opt_gubun, font=("Arial", 12, "bold"), width=28, state="readonly")
        combogubun['values'] = ('자산구매', '자산처분', '감가상각', '부채발생', '부채감소', '대표대여금', '세금계산서', '상품구매','기타원가항목','기타')
        txtvouno = Entry(entries_frame, textvariable=vouno, font=("Arial", 12), width=30)
        txtmemo = Entry(entries_frame, textvariable=memo, font=("Arial", 12, "bold"), width=30)

        txtpid.grid(row=0, column=1)
        txtlinesq.grid(row=1, column=1)
        txtsupplier.grid(row=2, column=1)
        txtindate.grid(row=3, column=1)
        txtvoudate.grid(row=4, column=1)
        comboproduct.grid(row=5, column=1)
        txtitem.grid(row=6, column=1)
        txtqty.grid(row=7, column=1)
        txtkg.grid(row=8, column=1)
        combocoa.grid(row=9, column=1)
        txtamount.grid(row=0, column=3, sticky="w")
        txtvat.grid(row=1, column=3, sticky="w")
        txtexrate.grid(row=2, column=3, sticky="w")
        txtcurrency.grid(row=3, column=3, sticky="w")
        txtdescript.grid(row=4, column=3, sticky="w")
        comboagroup.grid(row=5, column=3, sticky="w")
        combogubun.grid(row=6, column=3, sticky="w")
        txtvouno.grid(row=7, column=3, sticky="w")
        txtmemo.grid(row=8, column=3, sticky="w")

        # Fetch All Data from DB
        def fetch():
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            cursor.execute("SELECT * FROM cogs")
            rows = cursor.fetchall()
            rows=list(rows)
            rows.sort(key=lambda x: x[0], reverse=True)
            mydb.close()
            return rows

        def getData(event):
            selected_row = tv.focus()
            data = tv.item(selected_row)
            global row
            row = data["values"]
            #print(row)
            if len(row) >= 19: #assumming you have 19 columns
                pid.set(row[0])
                linesq.set(row[1])
                supplier.set(row[2])
                indate.set(row[3])
                voudate.set(row[4])
                comboproduct.set(row[5])
                item.set(row[6])
                qty.set(row[7])
                kg.set(row[8])
                combocoa.set(row[9])
                amount.set(row[10])
                vat.set(row[11])
                exrate.set(row[12])
                currency.set(row[13])
                descript.set(row[14])
                comboagroup.set(row[15])
                combogubun.set(row[16])
                vouno.set(row[17])
                memo.set(row[18])
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
                tv.column(col, width=max(70, max_width * 10))  # Minimum width of 70 pixels    

        def dispalyAll():
            tv.delete(*tv.get_children())
            for row in fetch():
                tv.insert("", END, values=row)
            adjustColumnWidths()

        def add_cog():
            
            pid_val = txtpid.get()
            linesq_val = txtlinesq.get()
            supplier_val = txtsupplier.get()
            indate_val = txtindate.get()
            voudate_val = txtvoudate.get()
            product_val = comboproduct.get()
            item_val = txtitem.get()
            qty_val = txtqty.get()
            kg_val = txtkg.get()
            coa_val = combocoa.get()
            amount_val = txtamount.get()
            vat_val = txtvat.get()
            exrate_val = txtexrate.get()
            currency_val = txtcurrency.get()
            descript_val = txtdescript.get()
            agroup_val = comboagroup.get()
            gubun_val = combogubun.get()
            vouno_val = txtvouno.get()
            memo_val = txtmemo.get()
            
            mydb = pymysql.connect(host="localhost", user="root", password="0000", database="fkl")
            cursor = mydb.cursor()
            
            if not (pid_val and voudate_val and coa_val and amount_val and agroup_val and gubun_val):
                messagebox.showerror('Error', 'Enter all fields.')
            else:
                try:
                    sql = "INSERT INTO  cogs (pid, linesq, supplier, indate, voudate, product, item, qty, kg, coa, amount, vat, exrate, currency, descript, agroup, gubun, vouno, memo)" \
                        "VALUES (%s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                    val = (pid_val, linesq_val, supplier_val, indate_val, voudate_val, product_val, item_val, qty_val, kg_val, coa_val, amount_val, vat_val, exrate_val, currency_val, descript_val, agroup_val, gubun_val, vouno_val, memo_val)
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

        def update_cog():
            selected_item = tv.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to update.')
            else:
                pid = txtpid.get()
                linesq = txtlinesq.get()
                supplier = txtsupplier.get()
                indate = txtindate.get()
                voudate = txtvoudate.get()
                product = comboproduct.get()
                item = txtitem.get()
                qty = txtqty.get()
                kg = txtkg.get()
                coa = combocoa.get()
                amount = txtamount.get()
                vat = txtvat.get()
                exrate = txtexrate.get()
                currency = txtcurrency.get()
                descript = txtdescript.get()
                agroup = comboagroup.get()
                gubun = combogubun.get()
                vouno = txtvouno.get()
                memo = txtmemo.get()
                
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()
            
                try:
                    sql = "UPDATE cogs set linesq=%s, supplier=%s, indate=%s,voudate=%s, product=%s, item=%s, qty=%s, kg=%s, coa=%s,amount=%s,vat=%s, exrate=%s,currency=%s," \
                            "descript=%s, agroup=%s, gubun=%s, vouno=%s, memo=%s where pid=%s"
                    val = (linesq, supplier, indate, voudate, product, item, qty, kg, coa, amount, vat, exrate, currency, descript, agroup, gubun, vouno, memo,pid)
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

        def delete_cog():
            selected_item = tv.focus()
            if not selected_item:
                messagebox.showerror('Error', 'Choose an row to delete.')
            else:
                pid = txtpid.get()
                mydb=pymysql.connect(host="localhost" , user="root" , password="0000", database="fkl")
                cursor=mydb.cursor()

                try:
                    sql = "DELETE from cogs where pid=%s"
                    val = (pid,)
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
            pid.set("")
            linesq.set("")
            supplier.set("")
            indate.set("")
            voudate.set("")
            comboproduct.set("")
            item.set("")
            qty.set("")
            kg.set("")
            combocoa.set("")
            amount.set("")
            vat.set("")
            exrate.set("")
            currency.set("")
            descript.set("")
            comboagroup.set("")
            combogubun.set("")
            vouno.set("")
            memo.set("")
            
        btn_frame = Frame(entries_frame, bg="powder blue")
        btn_frame.grid(row=10, column=0, columnspan=4, padx=10, pady=10, sticky="w")

        btnAdd = Button(btn_frame, command=add_cog, text="Add Details", width=15, font=("Calibri", 16, "bold"), fg="white",
                        bg="#16a085", bd=0).grid(row=0, column=0)
        btnEdit = Button(btn_frame, command=update_cog, text="Update Details", width=15, font=("Calibri", 16, "bold"),
                        fg="white", bg="#2980b9",
                        bd=0).grid(row=0, column=1, padx=10)
        btnDelete = Button(btn_frame, command=delete_cog, text="Delete Details", width=15, font=("Calibri", 16, "bold"),
                        fg="white", bg="#c0392b",
                        bd=0).grid(row=0, column=2, padx=10)
        btnClear = Button(btn_frame, command=clearAll, text="Clear Details", width=15, font=("Calibri", 16, "bold"), fg="white",
                        bg="#f39c12",
                        bd=0).grid(row=0, column=3, padx=10)

        #====Treeview Widget====
        tree_frame=LabelFrame(parent, text="COG DB LIST", bg="powder blue",bd=3, relief=RIDGE, font=("Arial", 12, "bold"),padx=10)
        tree_frame.place(x=0, y=510, width=1870, height=450)

        style = ttk.Style()
        style.configure("mystyle.Treeview", font=("Arial", 10),rowheight=30)  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=("Arial", 12, "bold"))  # Modify the font of the headings
        tv = ttk.Treeview(tree_frame, columns=(1, 2, 3, 4, 5, 6, 7, 8,9,10,11,12,13,14,15,16,17,18,19), style="mystyle.Treeview")
        tv.heading("1", text="ID")
        tv.heading("2", text="LineSq")
        tv.heading("3", text="Supplier")
        tv.heading("4", text="Indate")
        tv.heading("5", text="VouDate")
        tv.heading("6", text="Product")
        tv.heading("7", text="Item")
        tv.heading("8", text="Qty")
        tv.heading("9", text="Kg")
        tv.heading("10", text="COA")
        tv.heading("11", text="Amount")
        tv.heading("12", text="VAT")
        tv.heading("13", text="ExRate")
        tv.heading("14", text="Currency")
        tv.heading("15", text="Descript")
        tv.heading("16", text="Agroup")
        tv.heading("17", text="Gubun")
        tv.heading("18", text="Vouno")
        tv.heading("19", text="Memo")

        tv['show'] = 'headings'
        # tv.column("1", width=3)

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
#https://www.youtube.com/watch?v=7Z2J7NdsCRc 메뉴바 스타일힌트 사이트
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