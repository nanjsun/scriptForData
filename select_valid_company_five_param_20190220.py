import xlrd
import xlwt
import math
from xlrd import xldate_as_tuple


def select_valid_company_five_params():
    data = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190220\\两指标.xlsx')

    table = data.sheets()[0]  # 通过索引顺序获取

    valid_company_sheet_1998_2017 = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190219\\17-98.xlsx').sheets()[0]
    valid_company = get_valid_company(valid_company_sheet_1998_2017)
    book = xlwt.Workbook()  # 创建一个Excel
    res_sheet = book.add_sheet('valid_company_five_params')  # 在其中创建一个名为hello的sheet

    print(len(valid_company))
    rows = table.nrows
    # print(rows)
    clos = table.ncols
    res_sheet.write(0, 0, table.cell_value(0, 0))
    res_sheet.write(0, 1, table.cell_value(0, 1))
    res_sheet.write(0, 2, table.cell_value(0, 2))
    # res_sheet.write(0, 3, "所得税费用'")

    print("rows: " + str(rows))
    res_sheet_row_index = 1
    company_all = {}
    company_single = {}
    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[0]
        # year = int(table.row_values(i + 1)[1][0:4])
        current_time = xldate_as_tuple(table.row_values(i + 1)[1], 0)
        # print("year:" + str(current_time))
        # print("i:" + str(i))
        year = int(str(current_time[0]))
        param0 = table.cell_value(i + 1, 2)
        param1 = table.cell_value(i + 1, 3)
        param2 = table.cell_value(i + 1, 4)
        param3 = table.cell_value(i + 1, 5)
        param4 = table.cell_value(i + 1, 6)

        # tax_cost = table.cell_value(i + 1, 3)

        if valid_company.__contains__(stack_id) and 1997 < year < 2018:
            if company_all.__contains__(stack_id):
                company_all[stack_id][year] = [param0, param1, param2, param3, param4]
            else:
                company_single = {}
                company_single[year] = [param0, param1, param2, param3, param4]
                company_all[stack_id] = company_single
            # res_sheet.write(res_sheet_row_index, 0, stack_id)
            # res_sheet.write(res_sheet_row_index, 1, year)
            # res_sheet.write(res_sheet_row_index, 2, financial_cost)
            # res_sheet.write(res_sheet_row_index, 3, tax_cost)

    for stack_id in valid_company:
        for i in range(20):
            year = i + 1998
            res_sheet.write(res_sheet_row_index, 0, stack_id)
            res_sheet.write(res_sheet_row_index, 1, year)
            if company_all.__contains__(stack_id):
                if company_all[stack_id].__contains__(year):
                    if not company_all[stack_id][year] == "":
                        res_sheet.write(res_sheet_row_index, 2, company_all[stack_id][year][0])
                        res_sheet.write(res_sheet_row_index, 3, company_all[stack_id][year][1])
                        res_sheet.write(res_sheet_row_index, 4, company_all[stack_id][year][2])
                        res_sheet.write(res_sheet_row_index, 5, company_all[stack_id][year][3])
                        res_sheet.write(res_sheet_row_index, 6, company_all[stack_id][year][4])

                    else:
                        res_sheet.write(res_sheet_row_index, 2, 0)
                else:
                    res_sheet.write(res_sheet_row_index, 2, 0)
            else:
                res_sheet.write(res_sheet_row_index, 2, 0)
            res_sheet_row_index += 1

    book.save('C:\\Users\\nanj2\Documents\lulu\\20190220\\valid_company_interest_xxx_1998_2017.xls')


def get_valid_company(valid_company_sheet_1998_2017):
    nrows = valid_company_sheet_1998_2017.nrows

    valid_company = []

    for i in range(nrows - 1):
        company = valid_company_sheet_1998_2017.cell_value(i + 1, 0)
        if not valid_company.__contains__(company):
            valid_company.append(company)
    return valid_company


if __name__ == "__main__":
    select_valid_company_five_params()
