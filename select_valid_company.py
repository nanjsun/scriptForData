import xlrd
import xlwt
import math


def select_valid_company():
    data = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\all-1.xlsx')

    table = data.sheets()[0]  # 通过索引顺序获取

    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')  # 通过名称获取

    # 获取整行和整列的值（数组）
    # 　　
    # table.row_values(i)

    book = xlwt.Workbook()  # 创建一个Excel
    valid_company_sheet_1998_2017 = book.add_sheet('valid_company_1998_2017')  # 在其中创建一个名为hello的sheet

    rows = table.nrows
    # print(rows)
    clos = table.ncols

    print("rows: " + str(rows))
    for i in range(clos):
        valid_company_sheet_1998_2017.write(0, i, table.row_values(0)[i])

    first_year = {}
    last_year = {}
    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[0]
        current_time = table.row_values(i + 1)[1][0:4]
        money = table.row_values(i + 1)[2]
        current_year = (int)(current_time[0:4])
        # print("money:" + str(money))

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

        # if first_year[stack_id] != 1998:
        #     continue
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

        for j in range(clos) :
            valid_company_sheet_1998_2017.write(new_sheet_row_index, j, table.row_values(i + 1)[j])
        new_sheet_row_index += 1
        # print('hpindex: ', hp_index)

    print("rows of 1998 to 2017: " + str(new_sheet_row_index))

    book.save('C:\\Users\\nanj2\Documents\lulu\\20190214\\valid_company.xls')


if __name__ == "__main__":
    select_valid_company()
