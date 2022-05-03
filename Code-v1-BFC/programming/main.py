from programming import excel_function
from programming import column_calculator
from programming import well_position
from programming import data_save

def start(operator, client, order_number, order_name, step, column, order_date, due_date, file_name, parent_path):
    excel_function.create_excel(file_name, parent_path)
    excel_function.copy_step_format(file_name, step, parent_path)
    excel_function.basic_info_to_copied_format(operator, client, order_number, order_name, order_date, due_date, file_name, parent_path)
    column_height = column_calculator.column_calc(order_number)
    excel_function.volume_calculation(file_name, step, column_height, parent_path)
    well_position.main(file_name, order_number, column, parent_path)
    data_save.main(operator, client, order_number, order_name, step, order_date, due_date, file_name, parent_path)
