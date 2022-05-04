#created on: 2021 12 25
#Latest Update on: 2022 05 03
#author: github.com/F0RTRE55

import openpyxl
from tkinter import messagebox
from program import format_reader
from program import adder

def start(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name):
    #change format
    order_volume = float(order_volume)
    #create destination file
    format_reader.create_destination_file(current_directory, file_name)
    #create past order file
    format_reader.create_past_order_file(current_directory)
    #copy format to destination file
    format_reader.copy_file(current_directory, file_name, step)
    #add basic information to destination file
    adder.basic_information(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name)
    #calculate reactant position
    adder.position_calculation(current_directory, order_volume, step, file_name)
    #calculate reactant volume
    adder.volume_calculation(current_directory, order_volume, step, file_name)
    #calculate wells
    adder.well_calculation(current_directory, order_volume, well_new_or_old, wellID, well_position, file_name)
    #save order information to database
    adder.save_data_control(current_directory, operator, client, orderID, order_volume, well_new_or_old, wellID, well_position, step, order_date, deadline, file_name)