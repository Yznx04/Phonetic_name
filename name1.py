import xlrd
import win32com.client

workbook = xlrd.open_workbook("name.xlsx", "rb")
Data_sheet = workbook.sheets()[0]
name_list = Data_sheet.col_values(0)
s = 0
i = len(name_list)
while s < i:
    print("输入1开始点名")
    x = int(input())
    if x == 1:
        name = name_list[s]
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.rate = 1
        speaker.Speak(name)
        s += 1