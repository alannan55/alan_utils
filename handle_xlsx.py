from openpyxl import load_workbook

workbook = load_workbook(filename="E:\项目\诚善堂\付款信息\供应商货款\常用供应商付款信息.xlsx")

sheet = workbook['Sheet1']

cell_value = sheet['C1'].value

stop = 1