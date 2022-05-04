#created on: 2021 12 25
#Latest Update on: 2022 05 03
#author: github.com/F0RTRE55

import xlsxwriter
import openpyxl
from tkinter import messagebox 
from program import position
from program import volume
from program import well
from program import save_data

#add calculator inputs to destination file
def basic_information(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name):
    destination_directory = str(current_directory) + "\\result\\" + file_name + ".xlsx"
    workbook = openpyxl.load_workbook(destination_directory)
    worksheet = workbook.active
    worksheet = workbook.worksheets[0]
    #operator name
    worksheet['C3'] = operator
    #client name
    worksheet['H3'] = client
    #order volume
    worksheet['I8'] = order_volume
    #order name
    worksheet['D7'] = orderID
    #order date
    worksheet['C4'] = order_date
    #due date
    worksheet['H4'] = deadline
    #well ID
    worksheet['D8'] = wellID
    #save excel
    workbook.save(destination_directory)

#reactant position
def position_calculation(current_directory, order_volume, step, file_name):
    #go to appropriate step value for calculation
    if step == 1:
        position.step1(current_directory, order_volume, file_name)
    elif step ==2:
        position.step2(current_directory, order_volume, file_name)
    elif step ==3:
        position.step3(current_directory, order_volume, file_name)
    elif step ==4:
        position.step4(current_directory, order_volume, file_name)
    elif step ==5:
        position.step5(current_directory, order_volume, file_name)
    elif step ==6:
        position.step6(current_directory, order_volume, file_name)
    elif step ==7:
        position.step7(current_directory, order_volume, file_name)
    elif step ==8:
        position.step8(current_directory, order_volume, file_name)
    elif step ==9:
        position.step9(current_directory, order_volume, file_name)
    elif step ==10:
        position.step10(current_directory, order_volume, file_name)
    else:
        position.step11(current_directory, order_volume, file_name)
        
#reactant volume
def volume_calculation(current_directory, order_volume, step, file_name):
    #go to appropriate step value for calculation
    if step == 1:
        volume.step1(current_directory, order_volume, file_name)
    elif step ==2:
        volume.step2(current_directory, order_volume, file_name)
    elif step ==3:
        volume.step3(current_directory, order_volume, file_name)
    elif step ==4:
        volume.step4(current_directory, order_volume, file_name)
    elif step ==5:
        volume.step5(current_directory, order_volume, file_name)
    elif step ==6:
        volume.step6(current_directory, order_volume, file_name)
    elif step ==7:
        volume.step7(current_directory, order_volume, file_name)
    elif step ==8:
        volume.step8(current_directory, order_volume, file_name)
    elif step ==9:
        volume.step9(current_directory, order_volume, file_name)
    elif step ==10:
        volume.step10(current_directory, order_volume, file_name)
    else:
        volume.step11(current_directory, order_volume, file_name)

#well position
def well_calculation(current_directory, order_volume, well_new_or_old, wellID, well_position, file_name):
    if well_new_or_old == "new": #if well plate is new
        well.new_well(current_directory, order_volume, wellID, well_position, file_name)
    else: #if well plate is used
        well.used_well(current_directory, order_volume, wellID, well_position, file_name)

#save data except well
def save_data_control(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name):
    #ask question if user wants to save data
    reply = messagebox.askquestion("BFC V.4","Do you want to save order information?")
    if reply == "yes":
        empty_line = save_data.find_empty_line(current_directory)
        save_data.save_protocol(current_directory, empty_line, operator, client, order_volume, orderID, step, order_date, deadline, file_name)
    else: #will not save data
        exit()
