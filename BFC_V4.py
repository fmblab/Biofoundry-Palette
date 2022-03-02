import os, sys
import tkinter as tk
import tkinter.font as tkFont
from tkinter import *
from tkinter import messagebox
from tkcalendar import Calendar, DateEntry
from PIL import ImageTk, Image
from programming import main

#windows setup
root = tk.Tk()
root.title("BFC V.4")
root.geometry('260x300+40+80')
root.resizable(False, False)
root.configure(bg = 'white')

#frame icon
position = 'not'
if getattr(sys, 'position', False):
    position = 'ever'
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.getcwd() + "\images\Logo.ico"
root.iconbitmap(icon_path)

#Font Styles
TitleFontStyle = tkFont.Font(family = "Lucida Grande", size = 15, weight = "bold")
TextLabelStyle = tkFont.Font(family = "Lucida Grande", size = 10, weight = "bold")

#auto deletion
def operator_callback(event):
    operator.delete(0, "end")
    return None
def client_callback(event):
    client.delete(0, "end")
    return None
def order_number_callback(event):
    order_number.delete(0, "end")
    return None
def order_name_callback(event):
    order_name.delete(0, "end")
    return None
def file_name_callback(event):
    file_name.delete(0, "end")
    return None

#operator
tk.Label(root, text = "OPERATOR", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 0)
operator = tk.StringVar()
operator = tk.Entry(root, width = 20)
operator.place(x = 110,y = 0)
operator.insert(0, "ex: Min Gu Cho / 조민구")
operator.bind("<Button-1>", operator_callback)

#client
tk.Label(root, text = "CLIENT", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 30)
client = tk.StringVar()
client = tk.Entry(root, width = 20)
client.place(x = 110,y = 30)
client.insert(0, "ex: Min Gu Cho / 조민구")
client.bind("<Button-1>", client_callback)

#name of order
tk.Label(root, text = "주문 ID", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 60)
order_name = tk.IntVar()
order_name = tk.Entry(root, width = 20)
order_name.place(x = 110,y = 60)
order_name.insert(0, "ex: cg0007")
order_name.bind("<Button-1>", order_name_callback)

#number of orders
tk.Label(root, text = "주문 수", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 90)
order_number = tk.IntVar()
order_number = tk.Entry(root, width = 20)
order_number.place(x = 110,y = 90)
order_number.insert(0, "ex: 7 (max: 96)")
order_number.bind("<Button-1>", order_number_callback)

#step
tk.Label(root, text = "STEP", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 120)
step_var = tk.IntVar()
step_list = ['1','2','3','4','5','6','7','8','9','10','11']
step_spinbox = tk.Spinbox(root, values = step_list, textvariable = step_var, width = 5)
step_spinbox.place(x = 110, y = 120)

#column
tk.Label(root, text = "Well Position", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 150)
column_var = tk.IntVar()
column_list = ['1','2','3','4','5','6','7','8','9','10','11','12']
column_spinbox = tk.Spinbox(root, values = column_list, textvariable = column_var, width = 5)
column_spinbox.place(x = 110, y = 150)

#calendar
order_date_label = Label(root, text = "Order Date", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5, y = 180)
order_date = DateEntry(root, )
order_date.place(x = 110, y = 180)
due_date_label = Label(root, text = "Due Date", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 210)
due_date = DateEntry(root,)
due_date.place(x = 110,y = 210)

#file name
tk.Label(root, text = "파일명", bg = "#ffffff", fg = "#072b60", font = TextLabelStyle).place(x = 5,y = 240)
file_name = tk.StringVar()
file_name = tk.Entry(root, width = 20)
file_name.place(x = 110,y = 240)
file_name.insert(0, " ")
file_name.bind("<Button-1>", file_name_callback)

#confirmation
def onClick():
    tk.messagebox.showinfo("BFC V.4", "Program is Calculating")
    #get file path
    parent_path = str(os.getcwd())
    #operator
    g_operator = operator.get()
    #client
    g_client = client.get()
    #order number
    g_order_number = order_number.get()
    #order name
    g_order_name = order_name.get()
    #step
    g_step = step_spinbox.get()
    #column
    g_column = column_spinbox.get()
    #order date
    g_order_date = order_date.get()
    g_order_date = g_order_date.replace("/","-")
    #due date
    g_due_date = due_date.get()
    g_due_date = g_due_date.replace("/","-")
    #file name
    g_file_name = file_name.get()
    #send data to calculation
    main.start(g_operator, g_client, g_order_number, g_order_name, g_step, g_column, g_order_date, g_due_date, g_file_name, parent_path)
    
button_confirm = tk.Button(root, text = "Confirm", command = onClick, bg = "#8dc63f")
button_confirm.place(x = 101, y = 270)

root.mainloop()
