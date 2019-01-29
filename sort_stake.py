import xlrd
import xlwt


def sort_stake():
    origin = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190119\\1998-2018.xls')
    origin_table = origin.sheets()[0]
    # update_table = origin.add_sheet('new')
    origin_table_rows = origin_table.nrows
    origin_table_cols = origin_table.ncols
    # print(origin_table_rows)
    # print(origin_table_cols)
    new_book = xlwt.Workbook()  # 创建一个Excel
    update_table = new_book.add_sheet('sorted', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    for i in range(origin_table_cols):
        update_table.write(0, i, origin_table.row_values(0)[i])
    valid_stakes = []
    # print(origin_table_rows)
    for i in range(origin_table_rows - 1):
        stake_code = origin_table.row_values(i + 1)[0]
        if stake_code == "":
            break
        if not valid_stakes.__contains__(stake_code):
            valid_stakes.append(origin_table.row_values(i + 1)[0])
    valid_stakes.sort(key=None, reverse=False)
    # print(len(valid_stakes))
    print(valid_stakes)

    all_date = [{} for j in range(20)]
    # all_date = [dir() for j in range(20)]
    max_year = 0
    min_year = 10000
    for i in range(origin_table_rows - 1):
        stake_code = origin_table.row_values(i + 1)[0]
        year = int(origin_table.row_values(i + 1)[1])
        a = origin_table.row_values(i + 1)[2]
        b = origin_table.row_values(i + 1)[3]
        c = origin_table.row_values(i + 1)[4]
        d = origin_table.row_values(i + 1)[5]
        e = origin_table.row_values(i + 1)[6]
        f = origin_table.row_values(i + 1)[7]


        single_stake_directory = {stake_code: [year, a, b, c, d, e, f]}
        all_date[year - 1998][stake_code] = [stake_code, year, a, b, c, d, e, f]
        if year > max_year:
            max_year = year

        if year < min_year:
            min_year = year

    print(all_date[0])

    new_row_index = 0
    for valid_stake in valid_stakes:
        for i in range(max_year - min_year + 1):
            year = i + min_year
            year_data = all_date[year - 1998]
            a = ""
            b = ""
            c = ""
            d = ""
            e = ""
            f = ""
            if valid_stake in all_date[year - 1998]:
                # print(year_data[valid_stake])
                a = year_data[valid_stake][2]
                b = year_data[valid_stake][3]
                c = year_data[valid_stake][4]
                d = year_data[valid_stake][5]
                e = year_data[valid_stake][6]
                f = year_data[valid_stake][7]

            update_table.write(new_row_index, 0, valid_stake)
            update_table.write(new_row_index, 1, year)
            update_table.write(new_row_index, 2, a)
            update_table.write(new_row_index, 3, b)
            update_table.write(new_row_index, 4, c)
            update_table.write(new_row_index, 5, d)
            update_table.write(new_row_index, 6, e)
            update_table.write(new_row_index, 7, f)
            new_row_index += 1
    print(all_date[2017 - 1998].keys())
    new_book.save("C:\\Users\\nanj2\Documents\lulu\\20190119\\1998-2018-sorted.xls")


if __name__ == "__main__":
    sort_stake()


