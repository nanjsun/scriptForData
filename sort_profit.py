import xlrd
import xlwt


def sort_stake():
    origin = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190129\\profit-origin.xlsx')
    origin_table = origin.sheets()[0]
    # update_table = origin.add_sheet('new')
    origin_table_rows = origin_table.nrows
    origin_table_cols = origin_table.ncols
    # print(origin_table_rows)
    # print(origin_table_cols)
    new_book = xlwt.Workbook()  # 创建一个Excel
    update_table = new_book.add_sheet('sorted', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    # for i in range(origin_table_cols):
    #     update_table.write(0, i, origin_table.row_values(0)[i])
    valid_stakes = []
    # print(origin_table_rows)
    print(origin_table.row_values(0)[(2017 - 1998) * 2])
    print(origin_table.row_values(0)[(2017 - 1998) * 2 + 1])
    # print(origin_table.row_values(0)[(2017 - 1998) * 2 + 1])

    stake_code_year = origin_table.row_values(0)[(2017 - 1998) * 2]
    print("stake_code_year")
    print(stake_code_year)
    for i in range(origin_table_rows - 1):
        stake_code = origin_table.row_values(i + 1)[(2017 - 1998) * 2]
        if stake_code == "":
            break
        if not valid_stakes.__contains__(stake_code):
            valid_stakes.append(stake_code)
    valid_stakes.sort(key=None, reverse=False)
    print(valid_stakes[len(valid_stakes) -1])
    print(len(valid_stakes))
    # print(len(valid_stakes))
    print(valid_stakes)

    all_date = [{} for s in range(20)]
    max_year = 2017
    min_year = 1998
    for j in range(2017 - 1998 + 1):
        single_year = {}
        year = origin_table.row_values(0)[j * 2]

        for i in range(origin_table_rows - 1):
            stake_code = origin_table.row_values(i + 1)[j * 2]
            profit = origin_table.row_values(i + 1)[j * 2 + 1]
            single_year[stake_code] = profit
        all_date[int(year - 1998)] = single_year

    print(all_date[2012 - min_year].keys())
    # print(all_date[2012 - min_year]['600063'])

    update_table.write(0, 0, "Company")
    update_table.write(0, 1, "Year")
    update_table.write(0, 2, "Profit")
    new_row_index = 1
    for valid_stake in valid_stakes:
        for i in range(max_year - min_year + 1):
            year = i + min_year
            single_year = all_date[i]
            profit = 0
            if valid_stake in all_date[year - 1998]:
                # print(year_data[valid_stake])
                profit = single_year[valid_stake]

            update_table.write(new_row_index, 0, valid_stake)
            update_table.write(new_row_index, 1, year)
            update_table.write(new_row_index, 2, profit)

            new_row_index += 1
    # print(all_date[2017 - 1998].keys())
    new_book.save("C:\\Users\\nanj2\Documents\lulu\\20190129\\profit-sorted3.xls")


if __name__ == "__main__":
    sort_stake()


