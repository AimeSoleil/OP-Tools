# -*- coding: utf-8 -*-

import json
import datetime

YYYYMMDD = '%Y-%m-%d'
FILE_PATH = 'holidays_{0}.json'


class Holidays:
    def __init__(self, year):
        self.holidaysDict = self.loadHolidaysData(year)

    def isLeaveDay(self, date):
        if not any(self.holidaysDict):
            self.holidaysDict = self.loadHolidaysData(date.year)

        dateStr = date.strftime(YYYYMMDD)
        if self.holidaysDict.get(dateStr) and (self.holidaysDict.get(dateStr) == 1 or self.holidaysDict.get(dateStr) == 2):
            return True
        else:
            return False

    def loadHolidaysData(self, year):
        ''' http://www.easybots.cn/api/holiday.php?m=202001 '''
        holidaysFile = FILE_PATH.format(year)
        holidaysData = {}
        with open(holidaysFile, 'r', encoding='utf8') as holidays:
            jData = json.loads(holidays.read())
            for key in jData.keys():
                for key1 in jData[key].keys():
                    day = key[0:4] + '-' + key[4:] + '-' + key1
                    holidaysData[day] = int(jData[key][key1])

        return holidaysData


if __name__ == '__main__':
    d1 = datetime.datetime.strptime('2017-10-01', YYYYMMDD)
    holidays = Holidays(2017)
    print(holidays.isLeaveDay(d1))
