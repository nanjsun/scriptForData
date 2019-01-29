import xlrd
import xlwt


def remove_useless_stake():
    origin = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190117\\1998-2018-origin.xlsx')
    origin_table = origin.sheets()[0]
    # update_table = origin.add_sheet('new')
    origin_table_rows = origin_table.nrows
    origin_table_cols = origin_table.ncols
    # print(origin_table_rows)
    # print(origin_table_cols)
    new_book = xlwt.Workbook()  # 创建一个Excel
    update_table = new_book.add_sheet('new', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    for i in range(origin_table_cols):
        update_table.write(0, i, origin_table.row_values(0)[i])
    valid_stake = []
    print(origin_table_rows)
    for i in range(origin_table_rows - 1):
        if origin_table.row_values(i + 1)[0] == "":
            break
        valid_stake.append(origin_table.row_values(i + 1)[0])
    print(len(valid_stake))
    for i in range(int(origin_table_cols / 7)):
        new_row_index = 1
        for j in range(origin_table_rows - 1):
            stake = origin_table.row_values(j + 1)[i * 7]
            if not valid_stake.__contains__(stake):
                continue

            while not valid_stake[new_row_index - 1] == stake:
                new_row_index = new_row_index + 1
            for k in range(7):
                print(origin_table.row_values(j + 1)[i * 7 + k])
                update_table.write(new_row_index, i * 7 + k, origin_table.row_values(j + 1)[i * 7 + k])
            new_row_index = new_row_index + 1
    new_book.save("C:\\Users\\nanj2\Documents\lulu\\20190117\\new.xls")


if __name__ == "__main__":
    remove_useless_stake()


