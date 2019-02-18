import xlrd
import xlwt
import datetime
from xlrd import xldate_as_tuple


def mark_company_with_year():
    company_table = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190218\\放入.xlsx').sheets()[0]
    book2 = xlwt.Workbook()  # 创建一个Excel
    company_with_mark_sheet = book2.add_sheet('company_with_corruption', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    provinces_corruption = province_gdp()
    rows = company_table.nrows

    for i in range(rows - 1):
        company = company_table.cell_value(i + 1, 0)
        year = int(company_table.cell_value(i + 1, 1))
        province = company_table.cell_value(i + 1, 2)

        company_with_mark_sheet.write(i + 1, 0, company)
        company_with_mark_sheet.write(i + 1, 1, year)
        company_with_mark_sheet.write(i + 1, 2, province)
        company_with_mark_sheet.write(i + 1, 3, provinces_corruption[province][year])

    company_with_mark_sheet.write(0, 0, "company")
    company_with_mark_sheet.write(0, 1, "year")
    company_with_mark_sheet.write(0, 2, "province")
    company_with_mark_sheet.write(0, 3, "corruption")

    book2.save('C:\\Users\\nanj2\Documents\lulu\\20190218\\company_with_mark_province_corruption.xls')


def province_gdp():
    province_gdp_sheet = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190218\\立案人数.xlsx').sheets()[0]
    provinces_gdp = {}
    rows = province_gdp_sheet.nrows
    ncols = province_gdp_sheet.ncols

    for i in range(int(ncols - 1)):
        province = province_gdp_sheet.cell_value(0, i + 1)
        gdp = {}
        for j in range(rows - 1):
            gdp[1998 + j] = province_gdp_sheet.cell_value(j + 1, i + 1)

        provinces_gdp[province] = gdp
    return provinces_gdp


if __name__ == "__main__":
    # print(province_gdp())
    mark_company_with_year()
    # remove_redundant_year()
    # print(mark_province())