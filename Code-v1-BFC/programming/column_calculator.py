def column_calc (order_number):
    order_column = int(order_number) / 8
    if order_column <= 1:
        order_column = 1
    elif order_column <= 2:
        order_column = 2
    elif order_column <= 3:
        order_column = 3
    elif order_column <= 4:
        order_column = 4
    elif order_column <= 5:
        order_column = 5
    elif order_column <= 6:
        order_column = 6
    elif order_column <= 7:
        order_column = 7
    elif order_column <= 8:
        order_column = 8
    elif order_column <= 9:
        order_column = 9
    elif order_column <= 10:
        order_column = 10
    elif order_column <= 11:
        order_column = 11
    elif order_column <= 12:
        order_column = 12
    else:
        print("wrong input")
    return order_column
