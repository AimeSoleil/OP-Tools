# How to Generate exe file

> 需要在windows平台下，具体细节请参考：http://www.pyinstaller.org/

1. Python3 installed
2. Python -m pip install xlrd
3. Python -m pip install xlsxwriter
4. pyinstaller installed
5. pyinstaller --windowed --clean --noconfirm app.py

# 假期数据

目前的假期数据从这个API取过来的：`http://www.easybots.cn/api/holiday.php?m=202001`,

> 只更新到了2020，如果需要2021年的数据，请根据上面的API拉取数据，生成`holiday_data/holidays_2021.json` 文件，重新打包
> 如果上面的API访问不到，请自行构造数据

```json
  // 1 - 表示周末，2 - 表示法定假日
  "202003": {
    "01": "1",
    "07": "1",
    "08": "1",
    "14": "1",
    "15": "1",
    "21": "1",
    "22": "1",
    "28": "1",
    "29": "1"
  },
  "202004": {
    "04": "2",
    "05": "1",
    "06": "1",
    "11": "1",
    "12": "1",
    "18": "1",
    "19": "1",
    "25": "1"
  }
```

# 测试

`sample_data`目录下的Excel文件可以帮助测试
