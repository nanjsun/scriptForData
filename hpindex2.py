import xlrd
import xlwt
import math


def hp_index2():
    data = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190212\\AABB-origin.xlsx')

    table = data.sheets()[0]  # 通过索引顺序获取

    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')  # 通过名称获取

    # 获取整行和整列的值（数组）
    # 　　
    # table.row_values(i)

    book = xlwt.Workbook()  # 创建一个Excel
    hp_index_sheet = book.add_sheet('hpIndex')  # 在其中创建一个名为hello的sheet
    hp_index_sheet_1998_2017 = book.add_sheet('hpIndex_1998_2017')  # 在其中创建一个名为hello的sheet

    rows = table.nrows
    # print(rows)
    clos = table.ncols

    print("rows: " + str(rows))
    hp_index_sheet.write(0, 0, table.row_values(0)[0])
    hp_index_sheet.write(0, 1, table.row_values(0)[1])
    hp_index_sheet.write(0, 2, table.row_values(0)[2])
    hp_index_sheet.write(0, 3, table.row_values(0)[3])
    hp_index_sheet.write(0, 4, table.row_values(0)[4])
    hp_index_sheet.write(0, 5, table.row_values(0)[5])
    hp_index_sheet_1998_2017.write(0, 0, table.row_values(0)[0])
    hp_index_sheet_1998_2017.write(0, 1, table.row_values(0)[1])
    hp_index_sheet_1998_2017.write(0, 2, table.row_values(0)[2])
    hp_index_sheet_1998_2017.write(0, 3, table.row_values(0)[3])
    hp_index_sheet_1998_2017.write(0, 4, table.row_values(0)[4])
    hp_index_sheet_1998_2017.write(0, 5, table.row_values(0)[5])

    first_year = {}
    last_year = {}
    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[0]
        current_time = table.row_values(i + 1)[1][0:4]
        money = table.row_values(i + 1)[2]
        current_year = (int)(current_time[0:4])
        # print("money:" + str(money))
        if money > 0:
            ln_money = math.log(float(money) / (10**0))
        else:
            ln_money = 0
            print("money = 0 row: " + str(i + 1))

        if first_year.__contains__(stack_id):
            age = current_year - first_year.get(stack_id)
            if (first_year[stack_id]) > current_year:
                first_year[stack_id] = current_year
        else:
            age = 0
            first_year[stack_id] = current_year
        if last_year.__contains__(stack_id):
            if last_year[stack_id] < current_year:
                last_year[stack_id] = current_year
        else:
            last_year[stack_id] = current_year

        hp_index = -0.737 * ln_money + 0.043 * ln_money * ln_money - 0.04 * age
        # if first_year[stack_id] != 1998:
        #     continue
        hp_index_sheet.write(i + 1, 0, stack_id)
        hp_index_sheet.write(i + 1, 1, current_year)
        hp_index_sheet.write(i + 1, 2, money)
        hp_index_sheet.write(i + 1, 3, age)
        hp_index_sheet.write(i + 1, 4, ln_money)
        hp_index_sheet.write(i + 1, 5, hp_index)
        # print(age)
        # print('hpindex: ', hp_index)

    valid_keys = []
    print("first year size: " + str(len(first_year.keys())))
    print("last year size: " + str(len(last_year.keys())))
    for stack_id in first_year.keys():
        if first_year[stack_id] <= 1998 and last_year[stack_id] == 2017:
            valid_keys.append(stack_id)
        else:
            print(str(stack_id) + "first :" + str(first_year[stack_id]) + " last: " + str(last_year[stack_id]))

    print("valid keys size: " + str(len(valid_keys)))

    new_sheet_row_index = 1
    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[0]

        current_time = table.row_values(i + 1)[1][0:4]
        money = table.row_values(i + 1)[2]
        current_year = (int)(current_time[0:4])
        # print("money:" + str(money))
        if not valid_keys.__contains__(stack_id) or current_year < 1998:
            continue

        if money > 0:
            ln_money = math.log(float(money) / (10**0))
        else:
            ln_money = 0
            print("money = 0 row: " + str(i + 1))
        age = current_year - first_year[stack_id]

        hp_index = -0.737 * ln_money + 0.043 * ln_money * ln_money - 0.04 * age
        # if first_year[stack_id] != 1998:
        #     continue
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 0, stack_id)
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 1, current_year)
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 2, money)
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 3, age)
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 4, ln_money)
        hp_index_sheet_1998_2017.write(new_sheet_row_index, 5, hp_index)
        new_sheet_row_index += 1
        # print('hpindex: ', hp_index)

    print("rows of 1998 to 2017: " + str(new_sheet_row_index))

    book.save('C:\\Users\\nanj2\Documents\lulu\\20190212\\AABB-hp-2.xls')


if __name__ == "__main__":
    hp_index2()
