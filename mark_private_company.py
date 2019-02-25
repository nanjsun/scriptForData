import xlrd
import xlwt
import datetime
from xlrd import xldate_as_tuple


def mark_private_company():
    company_table = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190223\\2.xlsx').sheets()[0]
    result_table = xlwt.Workbook()  # 创建一个Excel
    company_with_private_mark_sheet = result_table.add_sheet('company_with_private_mark', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    private_company_list = get_private_company_list()

    rows = company_table.nrows

    for i in range(rows - 1):
        stark_id = company_table.cell_value(i + 1, 0)
        year = company_table.cell_value(i + 1, 1)
        private = 0

        if private_company_list.__contains__(stark_id):
            years = private_company_list[stark_id]
            if years.__contains__(year):
                private = 1

        company_with_private_mark_sheet.write(i + 1, 0, stark_id)
        company_with_private_mark_sheet.write(i + 1, 1, year)
        company_with_private_mark_sheet.write(i + 1, 2, private)

    company_with_private_mark_sheet.write(0, 0, "company")
    company_with_private_mark_sheet.write(0, 1, "year")
    company_with_private_mark_sheet.write(0, 2, "private")
    result_table.save('C:\\Users\\nanj2\Documents\lulu\\20190223\\company_with_private_mark.xls')


def get_private_company_list():
    private_company_excel = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190223\\民企.xlsx')
    table = private_company_excel.sheets()[0]  # 通过索引顺序获取
    # book = xlwt.Workbook()  # 创建一个Excel
    # company_with_single_year = book.add_sheet('company_with_single_year')  # 在其中创建一个名为hello的sheet
    rows = table.nrows
    # cols = table.ncols
    print("rows----lulu: " + str(rows) + "个")
    print(rows + 10)

    # first_year = {}
    private_company_list = {}
    years = []
    for i in range(rows - 1):
        stack_id = table.cell_value(i + 1, 0)
        year = table.cell_value(i + 1, 2)

        if private_company_list.__contains__(stack_id):
            years = private_company_list[stack_id]
            years.add(year)
        else:
            years = {year}
        private_company_list[stack_id] = years
    return private_company_list


if __name__ == "__main__":
    # print(province_gdp())
    # mark_company_with_year()
    mark_private_company()
    # get_private_company_list()
    # remove_redundant_year()
    # print(mark_province())