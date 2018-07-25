# -*- coding: utf-8 -*-
import os
import csv
import xlrd
import xlsxwriter
import datetime
from datetime import timedelta
from utils import Holidays

DATE_FORMAT_YYYY_MM_dd = '%Y-%m-%d'
DATE_FORMAT = '%Y-%m-%d %H:%M'
START_DATE_FORMAT = '%Y-%m-%d 09:00'
END_DATE_FORMAT = '%Y-%m-%d 18:00'
NOON_DATE_FORMAT = '%Y-%m-%d 12:00'
AFTERNOON_DATE_FORMAT = '%Y-%m-%d 13:00'
BG_CF_INDEX = 0
BU_INDEX = 1
WORK_LOCATION_INDEX = 2
BUSINESS_INDEX = 3
DEPARTMENT_INDEX = 4
EMPLOYEE_ID_INDEX = 5
NAME_INDEX = 6
APPLY_FOR_TIME_INDEX = 7
LEAVE_TYPE_INDEX = 8
START_DATE_INDEX = 9
END_DATE_INDEX = 10
HOUR_DELTA_INDEX = 11
STATE_INDEX = 12
APPROVER_INDEX = 13
LEAVE_TYPE_OF_SPECIAL = '其他'


# input xlsx to csv
def csv_from_excel(input):
    wb = xlrd.open_workbook(input)
    sh_list = wb.sheet_names()
    sh = wb.sheet_by_name(sh_list[0])
    temp_csv = os.path.join(os.getcwd(), "temp.csv")
    with open(temp_csv, mode="w", encoding="utf-8") as temp:
        wr = csv.writer(temp, quoting=csv.QUOTE_MINIMAL)
        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))
    return temp_csv


# input csv to xlsx
def csv_to_excel(input_csv):
    output_xlsx = input_csv.replace(".csv", ".xlsx")
    wb = xlsxwriter.Workbook(output_xlsx)
    ws = wb.add_worksheet("Sheet 1")
    with open(input_csv, mode="r") as csvfile:
        table = csv.reader(csvfile)
        i = 0
        for row in table:
            write_worksheet_row(wb, ws, row, i)
            i += 1
    wb.close()
    return output_xlsx


def write_worksheet_row(wb, ws, row, row_index):
    ws.write_string(row_index, 0, row[BG_CF_INDEX])
    ws.write_string(row_index, 1, row[BU_INDEX])
    ws.write_string(row_index, 2, row[WORK_LOCATION_INDEX])
    ws.write_string(row_index, 3, row[DEPARTMENT_INDEX])

    if row_index == 0:
        ws.write_string(row_index, 4, row[APPLY_FOR_TIME_INDEX])
    else:
        cell_date_format = wb.add_format({'num_format': 'yyyy/mm/dd hh:mm', 'align': 'left'})
        ws.write_datetime(row_index, 4, strToDate(row[APPLY_FOR_TIME_INDEX], DATE_FORMAT), cell_date_format)

    ws.write_string(row_index, 5, row[EMPLOYEE_ID_INDEX])

    if row_index == 0:
        ws.write_string(row_index, 6, row[START_DATE_INDEX])
    else:
        cell_date_format = wb.add_format({'num_format': 'yyyy/mm/dd', 'align': 'left'})
        ws.write_datetime(row_index, 6, strToDate(row[START_DATE_INDEX][0:10], DATE_FORMAT_YYYY_MM_dd), cell_date_format)

    # if row_index == 0:
    #     ws.write_string(row_index, 7, row[END_DATE_INDEX])
    # else:
    #     if (strToDate(row[START_DATE_INDEX], DATE_FORMAT) - strToDate(row[END_DATE_INDEX], DATE_FORMAT)).seconds != 0:
    #         cell_date_format = wb.add_format({'num_format': 'yyyy/mm/dd hh:mm', 'align': 'left'})
    #         ws.write_datetime(row_index, 7, strToDate(row[END_DATE_INDEX], DATE_FORMAT), cell_date_format)

    ws.write_string(row_index, 7, row[LEAVE_TYPE_INDEX])

    if row_index == 0:
        ws.write_string(row_index, 10, row[HOUR_DELTA_INDEX])
    else:
        ws.write_number(row_index, 10, float(row[HOUR_DELTA_INDEX]))

    ws.write_string(row_index, 11, row[BUSINESS_INDEX])
    ws.write_string(row_index, 12, row[NAME_INDEX])
    ws.write_string(row_index, 13, row[STATE_INDEX])
    ws.write_string(row_index, 14, row[APPROVER_INDEX])
    if row_index == 0:
        ws.write_string(row_index, 8, "Project Code")
        ws.write_string(row_index, 9, "Content")
    else:
        ws.write_string(row_index, 8, "")
        ws.write_string(row_index, 9, "")



# input csv, output csv
def process(input, temp_output):
    year = datetime.datetime.today().year
    holidays = Holidays(year)
    with open(input, 'r', encoding='utf8') as in_csv, open(temp_output, 'w', newline='') as out_csv:
        writer = csv.writer(out_csv, quotechar='"', quoting=csv.QUOTE_MINIMAL)
        reader = csv.reader(in_csv)

        o_index = 0
        for row in reader:
            o_index += 1
            if o_index <= 1:
                writer.writerow(row[:])  # 添加头部
            else:
                if not row:    # 如果list为空直接跳过
                    continue

                startDateStr = resolveEncoding(row[START_DATE_INDEX])
                endDateStr = resolveEncoding(row[END_DATE_INDEX])

                startDate = strToDate(startDateStr, DATE_FORMAT)
                endDate = strToDate(endDateStr, DATE_FORMAT)

                plusADay = timedelta(days=+1)

                if (endDate - startDate).days == 0:
                    if not holidays.isLeaveDay(startDate):
                        resolveDateRangeSmallerThan8Hs(startDate, endDate, row, writer, holidays)   # 如果相隔不超过一天，直接添加数据
                else:                      # 否则需要将date range拆分
                    inner_index = 0
                    newRow = row[:]
                    while (endDate - startDate).days >= 0:
                        startOfWorkDay = strToDate(dateToStr(startDate, START_DATE_FORMAT), DATE_FORMAT)
                        if inner_index == 0:
                            if (startDate - startOfWorkDay).seconds > 0:
                                row_to_fill = newRow[:]
                                new_startDate_1 = startOfWorkDay
                                new_endDate_1 = startDate
                                row_to_fill[START_DATE_INDEX] = dateToStr(new_startDate_1, DATE_FORMAT)
                                row_to_fill[END_DATE_INDEX] = dateToStr(new_endDate_1, DATE_FORMAT)
                                row_to_fill[HOUR_DELTA_INDEX] = calDiffDayHours(new_startDate_1, new_endDate_1)
                                row_to_fill[LEAVE_TYPE_INDEX] = LEAVE_TYPE_OF_SPECIAL
                                if not holidays.isLeaveDay(new_startDate_1) and calDiffDayHours(new_startDate_1, new_endDate_1) <= 8:
                                    writer.writerow(row_to_fill)

                            newRow[START_DATE_INDEX] = dateToStr(startDate, DATE_FORMAT)
                            org_startDate = startDate

                            startDate = startDate + plusADay
                        else:
                            newRow[START_DATE_INDEX] = dateToStr(startDate, START_DATE_FORMAT)
                            org_startDate = strToDate(dateToStr(startDate, START_DATE_FORMAT), DATE_FORMAT)

                            startDate = strToDate(dateToStr(startDate + plusADay, START_DATE_FORMAT), DATE_FORMAT)
                        inner_index += 1

                        new_endDate = startDate
                        if (new_endDate - endDate).days >= 0:
                            new_endDate = endDate
                            newRow[END_DATE_INDEX] = dateToStr(new_endDate, DATE_FORMAT)
                            newRow[HOUR_DELTA_INDEX] = calDiffDayHours(org_startDate, new_endDate)

                            if not holidays.isLeaveDay(new_endDate) and calDiffDayHours(org_startDate, new_endDate) <= 8:
                                writer.writerow(newRow)

                            # 如果new_endDate is before 18:00
                            endOfThisDay = strToDate(dateToStr(new_endDate, END_DATE_FORMAT), DATE_FORMAT)
                            if (endOfThisDay - new_endDate).seconds > 0:
                                new_startDate = new_endDate
                                newRow1 = newRow[:]
                                newRow1[START_DATE_INDEX] = dateToStr(new_startDate, DATE_FORMAT)
                                newRow1[END_DATE_INDEX] = dateToStr(endOfThisDay, DATE_FORMAT)
                                newRow1[HOUR_DELTA_INDEX] = calDiffDayHours(new_startDate, endOfThisDay)
                                newRow1[LEAVE_TYPE_INDEX] = LEAVE_TYPE_OF_SPECIAL

                                if not holidays.isLeaveDay(new_startDate) and calDiffDayHours(new_startDate, endOfThisDay) <= 8:
                                    writer.writerow(newRow1)
                        else:
                            new_endDate = strToDate(dateToStr(org_startDate, END_DATE_FORMAT), DATE_FORMAT)
                            newRow[END_DATE_INDEX] = dateToStr(new_endDate, DATE_FORMAT)
                            newRow[HOUR_DELTA_INDEX] = calDiffDayHours(org_startDate, new_endDate)

                            if not holidays.isLeaveDay(new_endDate) and calDiffDayHours(org_startDate, new_endDate) <= 8:
                                writer.writerow(newRow)

    os.remove(input)
    result_xlsx = csv_to_excel(temp_output)
    os.remove(temp_output)

    return result_xlsx


def resolveDateRangeSmallerThan8Hs(startDate, endDate, row, writer, holidays):
    resolveDateRange1(startDate, row, writer, holidays)
    if not holidays.isLeaveDay(startDate):
        writer.writerow(row[:])
    resolveDateRange2(endDate, row, writer, holidays)


def resolveDateRange1(startDate, row, writer, holidays):
    startOfWorkDay = strToDate(dateToStr(startDate, START_DATE_FORMAT), DATE_FORMAT)
    if (startDate - startOfWorkDay).seconds > 0:
        newRow1 = row[:]
        new_startDate = startOfWorkDay
        new_endDate = startDate
        newRow1[START_DATE_INDEX] = dateToStr(new_startDate, DATE_FORMAT)
        newRow1[END_DATE_INDEX] = dateToStr(new_endDate, DATE_FORMAT)
        newRow1[HOUR_DELTA_INDEX] = calDiffDayHours(new_startDate, new_endDate)
        newRow1[LEAVE_TYPE_INDEX] = LEAVE_TYPE_OF_SPECIAL
        if not holidays.isLeaveDay(new_startDate) and calDiffDayHours(new_startDate, new_endDate) <= 8:
            writer.writerow(newRow1)


def resolveDateRange2(endDate, row, writer, holidays):
    endOfWorkDay = strToDate(dateToStr(endDate, END_DATE_FORMAT), DATE_FORMAT)

    if (endOfWorkDay - endDate).seconds > 0:
        newRow2 = row[:]
        new_startDate = endDate
        new_endDate = endOfWorkDay
        newRow2[START_DATE_INDEX] = dateToStr(new_startDate, DATE_FORMAT)
        newRow2[END_DATE_INDEX] = dateToStr(new_endDate, DATE_FORMAT)
        newRow2[HOUR_DELTA_INDEX] = calDiffDayHours(new_startDate, new_endDate)
        newRow2[LEAVE_TYPE_INDEX] = LEAVE_TYPE_OF_SPECIAL
        if not holidays.isLeaveDay(new_startDate) and calDiffDayHours(new_startDate, new_endDate) <= 8:
            writer.writerow(newRow2)


def resolveEncoding(value):
    bytes_value = u''.join(value).encode('utf-8').strip()
    return bytes_value.decode('utf-8')


# return datetime
def strToDate(dateStr, format):
    return datetime.datetime.strptime(dateStr, format)


# return date string: 2017-06-29 15:00
def dateToStr(date, format):
    return date.strftime(format)


# calculate hour between startDate and endDate
#     -0      case1           s,e
#           case2           s                  e
#           case3           s                                     e
#           case4                             s,e
#           case5                             s                   e
#     -0     case6                                                 s,e
#    ==========|09:00|================|12:00|=====|13:00|========================|18:00|=======
def calDiffDayHours(startDate, endDate):
    noon_1200 = strToDate(dateToStr(startDate, NOON_DATE_FORMAT), DATE_FORMAT)
    afternoon_1300 = strToDate(dateToStr(startDate, AFTERNOON_DATE_FORMAT), DATE_FORMAT)
    startOfWorkDay = strToDate(dateToStr(startDate, START_DATE_FORMAT), DATE_FORMAT)
    endOfWorkDay = strToDate(dateToStr(startDate, END_DATE_FORMAT), DATE_FORMAT)

    if startDate < startOfWorkDay:
        startDate = startOfWorkDay
    if endDate > endOfWorkDay:
        endDate = endOfWorkDay

    if endDate <= noon_1200 or startDate >= afternoon_1300:
        # case1 & case6
        return (endDate - startDate).seconds / 3600
    if startDate < noon_1200 and (endDate >= noon_1200 and endDate <= afternoon_1300):
        # case2
        return (noon_1200 - startDate).seconds / 3600
    if startDate < noon_1200 and endDate > afternoon_1300:
        # case3
        return ((noon_1200 - startDate).seconds / 3600) + ((endDate - afternoon_1300).seconds / 3600)
    if (endDate >= noon_1200 and endDate <= afternoon_1300) and (startDate >= noon_1200 and startDate <= afternoon_1300):
        # case4
        return 0
    if (startDate >= noon_1200 and startDate <= afternoon_1300) and endDate >= afternoon_1300:
        # case5
        return (endDate - afternoon_1300).seconds / 3600


if __name__ == '__main__':
    # input = 'C:\\Users\\pact\\Desktop\\Marina_TimeSheet_Data\\marina_timesheet_input.csv'
    # output = 'C:\\Users\\pact\\Desktop\\Marina_TimeSheet_Data\\marina_timesheet_output.csv'
    input = 'C:\\Users\Administrator\Desktop\marina_timesheet_output.xlsx'
    output = 'C:\\Users\Administrator\Desktop\output.csv'

    process(csv_from_excel(input), output)
