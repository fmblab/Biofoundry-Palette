#created on: 2021 12 25
#Latest Update on: 2022 05 03
#author: github.com/F0RTRE55

from dataclasses import replace
import os #for finding .exe directory
from tkinter import * #for program UI
from tkinter import messagebox #for alert messages
import tkinter.font as tkfont #for UI font
from turtle import bgcolor, onclick
from matplotlib.pyplot import step #for program UI
from tkcalendar import Calendar, DateEntry #for order and deadline calendar
from program import bridge

#.exe directory search
directory = os.getcwd()

#UI frame codes
root = Tk() #program UI root
root.title("Biofoundry Palette") #UI title
root.geometry("280x250+40+80") #UI size
root.resizable(TRUE,False) #fixed UI size
root.configure(bg="white") #UI background color
root.iconbitmap(directory+"\program\logo.ico") #UI logo

#UI font code
fontstyle = tkfont.Font(family = "Lucida Grande", size = 10, weight = "bold")

#auto delete placeholder when input box is clicked codes
def operator_placeholder_delete(event):
    operator_input.configure(state=NORMAL)
    operator_input.delete(0,END) #delete all operator input placeholder content
    return None
def client_placeholder_delete(event):
    client_input.configure(state=NORMAL)
    client_input.delete(0,END) #delete all client input placeholder content
    return None
def orderID_placeholder_delete(event):
    orderID_input.configure(state=NORMAL)
    orderID_input.delete(0,END) #delete all orderID input placeholder content
    return None
def order_volume_placeholder_delete(event):
    order_volume_input.configure(state=NORMAL)
    order_volume_input.delete(0,END) #delete all order volume input placeholder content
    return None

#operator codes
operator_label = Label(root, text="OPERATOR:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #operator label
operator_label.grid(row = 0, column = 0) #operator label grid
operator_input = Entry(root, width = 20) #operator input
operator_input.grid(row = 0, column = 1) #operator input grid
operator_input.insert(0, "ex: Min Gu Cho") #operator input placeholder
operator_input.bind("<Button-1>",operator_placeholder_delete) #when button clicked delete operator input placeholder

#client codes
client_label = Label(root, text="CLIENT:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #client label
client_label.grid(row = 1, column = 0) #client label grid
client_input = Entry(root, width = 20) #client input
client_input.grid(row = 1, column = 1) #client input grid
client_input.insert(0, "ex: Min Gu Cho") #client input placeholder
client_input.bind("<Button-1>",client_placeholder_delete) #when button clicked delete client input placeholder

#order ID codes
orderID_label = Label(root, text="ORDER ID:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #orderID label
orderID_label.grid(row = 2, column = 0) #orderID label grid
orderID_input = Entry(root, width = 20) #orderID input
orderID_input.grid(row = 2, column = 1) #orderID input grid
orderID_input.insert(0, "ex: cg0007") #orderID input placeholder
orderID_input.bind("<Button-1>",orderID_placeholder_delete) #when button clicked delete orderID input placeholder

#order volume codes
order_volume_label = Label(root, text="ORDER VOLUME:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #order volume label
order_volume_label.grid(row = 3, column = 0) #order volume label grid
order_volume_input = Entry(root, width = 20) #order volume input
order_volume_input.grid(row = 3, column = 1) #order volume input grid
order_volume_input.insert(0, "ex: 7 (max: 96)") #order volume input placeholder
order_volume_input.bind("<Button-1>",order_volume_placeholder_delete) #when button clicked delete order volume input placeholder

#well codes
well_new_or_old_var = StringVar() #group radio button and value is set to int
well_new_or_old_var.set("new") #set well new or old default to new
well_id_label = Label(root, text="WELL ID:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #well id label
well_id_label.grid(row = 5, column = 0) #well id label grid
well_id_input = Entry(root, width = 20) #well id input
well_id_input.grid(row = 5, column = 1) #well id input grid
well_position_label = Label(root, text="WELL POSITION:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #well position label
well_position_label.grid(row = 6, column = 0) #well position label grid
well_position_input = IntVar()
well_position_list = ['1','2','3','4','5','6','7','8','9','10','11','12'] #available well column lists
well_position_spinbox = Spinbox(root, values= well_position_list, textvariable = well_position_input, width=5) #well starting position spinbox
well_position_spinbox.grid(row = 6, column = 1) #well position input grid

#step codes
step_label = Label(root, text="STEP:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #step label
step_label.grid(row = 7, column = 0) #step label grid
step_input = IntVar()
step_list = ['1','2','3','4','5','6','7','8','9','10','11'] #available step lists
step_spinbox = Spinbox(root, values= step_list, textvariable = step_input, width=5) #step spinbox
step_spinbox.grid(row = 7, column = 1) #step spinbox grid

#order and deadline dates
order_date_label = Label(root, text = "ORDER DATE:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #order date label
order_date_label.grid(row = 8, column = 0) #order date label grid
order_date_calendar = DateEntry(root, ) #order date calendar
order_date_calendar.grid(row = 8, column = 1) #order date calendar grid
deadline_label = Label(root, text = "DEADLINE:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #deadline label
deadline_label.grid(row = 9, column = 0) #deadline label grid
deadline_calendar = DateEntry(root,) #deadline calendar
deadline_calendar.grid(row = 9, column = 1) #deadline calendar grid

#Destination File Name
file_name_label = Label(root, text = "FILE NAME:", font = fontstyle, bg = "#ffffff", fg = "#072b60", anchor= "e", width = "15") #file name label
file_name_label.grid(row=10, column=0) #file name label grid
file_name_input = Entry(root, width = 20) #file name input label
file_name_input.grid(row=10,column=1) #file name input grid

#confirmation button event codes
def onClick():
    messagebox.showinfo("BFC","Program is Calculating")
    #gather inputted data
    current_directory = directory
    operator = operator_input.get()
    client = client_input.get()
    orderID = orderID_input.get()
    order_volume = order_volume_input.get()
    well_new_or_old = well_new_or_old_var.get()
    wellID = well_id_input.get()
    well_position = well_position_input.get()
    step = step_input.get()
    order_date = order_date_calendar.get()
    order_date = order_date.replace(". ","-") #change date formats
    order_date = order_date.replace(".","")
    deadline = deadline_calendar.get()
    deadline = deadline.replace(". ","-")
    deadline = deadline.replace(".","")
    file_name = file_name_input.get()
    #send gathered data to bridge.py
    bridge.start(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name)

confirm_button = Button(root, text="CALCULATE", command = onClick, bg = "#8dc63f")
confirm_button.grid(row=11,column=1)

root.mainloop() #program UI loop
