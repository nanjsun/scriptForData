import xlrd
import xlwt
import datetime
from xlrd import xldate_as_tuple


def mark_company_with_year():
    company_table = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\111.xlsx').sheets()[0]
    book2 = xlwt.Workbook()  # 创建一个Excel
    company_with_mark_sheet = book2.add_sheet('company_with_mark', cell_overwrite_ok=True)  # 在其中创建一个名为hello的sheet
    company_mark_year = get_mark_year()
    company_provinces = mark_province()
    company_industry = mark_industry()
    provinces_gdp = province_gdp()
    provinces_corruption = province_corruption()
    print(company_provinces.keys())
    rows = company_table.nrows

    for i in range(rows - 1):
        company = company_table.cell_value(i + 1, 0)
        year = int(company_table.cell_value(i + 1, 1))
        province = company_provinces[company]

        company_with_mark_sheet.write(i + 1, 0, company)
        company_with_mark_sheet.write(i + 1, 1, year)
        company_with_mark_sheet.write(i + 1, 3, province)
        company_with_mark_sheet.write(i + 1, 5, provinces_gdp[province][year])
        company_with_mark_sheet.write(i + 1, 6, provinces_corruption[province][year])
        if company_industry.__contains__(company):
            company_with_mark_sheet.write(i + 1, 4, company_industry[company])
        else:
            company_with_mark_sheet.write(i + 1, 4, "none")
            print("no industry: + " + company)

        if not company_mark_year.__contains__(company):
            print("without company: " + company)
            company_with_mark_sheet.write(i + 1, 2, 0)

            continue
        if year < company_mark_year[company]:
            company_with_mark_sheet.write(i + 1, 2, 0)
        else:
            company_with_mark_sheet.write(i + 1, 2, 1)
    company_with_mark_sheet.write(0, 0, "company")
    company_with_mark_sheet.write(0, 1, "year")
    company_with_mark_sheet.write(0, 2, "mark")
    company_with_mark_sheet.write(0, 3, "province")
    company_with_mark_sheet.write(0, 4, "industry")
    company_with_mark_sheet.write(0, 5, "province gdp")
    company_with_mark_sheet.write(0, 5, "province corruption")
    book2.save('C:\\Users\\nanj2\Documents\lulu\\20190214\\company_with_mark_province_industry_gdp_corruption.xls')


def province_gdp():
    province_gdp_sheet = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\各省GDP.xlsx').sheets()[0]
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


def province_corruption():
    province_gdp_sheet = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\各省腐败.xlsx').sheets()[0]
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


def mark_province():
    province_sheet = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\各省企业.xlsx').sheets()[0]
    company_provinces = {}

    rows = province_sheet.nrows
    ncols = province_sheet.ncols

    for i in range(int(ncols / 2)):
        for j in range(rows - 1):
            company = province_sheet.cell_value(j + 1, i * 2 + 0)
            province = province_sheet.cell_value(j + 1, i * 2 + 1)
            company_provinces[company] = province
    return company_provinces


def mark_industry():
    province_sheet = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\总行业.xlsx').sheets()[0]
    company_provinces = {}

    rows = province_sheet.nrows
    ncols = province_sheet.ncols

    for i in range(int(ncols / 2)):
        for j in range(rows - 1):
            company = province_sheet.cell_value(j + 1, i * 2 + 0)
            province = province_sheet.cell_value(j + 1, i * 2 + 1)
            company_provinces[company] = province
    return company_provinces


def get_mark_year():
    data = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20190214\\state_company.xlsx')
    table = data.sheets()[0]  # 通过索引顺序获取
    book = xlwt.Workbook()  # 创建一个Excel
    company_with_single_year = book.add_sheet('company_with_single_year')  # 在其中创建一个名为hello的sheet
    rows = table.nrows
    cols = table.ncols
    print("rows: " + str(rows))
    for i in range(2):
        company_with_single_year.write(0, i, table.row_values(0)[i])
    first_year = {}
    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[0]
        if stack_id == "":
            break
        current_time = xldate_as_tuple(table.row_values(i + 1)[1], 0)
        # print("year:" + str(current_time))
        # print("i:" + str(i))
        current_year = int(str(current_time[0]))
        # print("money:" + str(money))
        if first_year.__contains__(stack_id):
            print(str(stack_id) + ":" + str(first_year[stack_id]) + ":" + str(current_year))

            if (first_year[stack_id]) > current_year:
                first_year[stack_id] = current_year
        else:
            first_year[stack_id] = current_year

        # if first_year[stack_id] != 1998:
        #     continue
        # print(age)
        # print('hpindex: ', hp_index)

    for i in range(rows - 1):
        stack_id = table.row_values(i + 1)[2]
        if stack_id == "":
            break
        current_time = xldate_as_tuple(table.row_values(i + 1)[1], 0)
        # print("year:" + str(current_time))
        # print("i:" + str(i))
        current_year = int(str(current_time[0]))

        # print("money:" + str(money))
        if first_year.__contains__(stack_id):
            print(str(stack_id) + ":" + str(first_year[stack_id]) + ":" + str(current_year))
            if (first_year[stack_id]) > current_year:
                first_year[stack_id] = current_year
        else:
            first_year[stack_id] = current_year

        # if first_year[stack_id] != 1998:
        #     continue
        # print(age)
        # print('hpindex: ', hp_index)

    new_sheet_row_index = 0
    for stack_id in first_year.keys():
        company_with_single_year.write(new_sheet_row_index + 1, 0, stack_id)
        company_with_single_year.write(new_sheet_row_index + 1, 1, first_year[stack_id])
        new_sheet_row_index += 1

    # book.save('C:\\Users\\nanj2\Documents\lulu\\20190214\\state_company_single_year.xls')
    return first_year


if __name__ == "__main__":
    # print(province_gdp())
    mark_company_with_year()
    # remove_redundant_year()
    # print(mark_province())