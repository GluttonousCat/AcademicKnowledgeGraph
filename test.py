import win32api
from main import excelLineChart, writeIntoExcel

writeIntoExcel('test', 'test', [{'x1':1, 'y1':2, 'z1':3}, {'x1':1, 'y1':2, 'z1':3}])
win32api.ShellExecute(0, 'open', 'test.xlsx', '', '', 1)