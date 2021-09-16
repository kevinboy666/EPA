import win32com.client
import os
import wx

wxapp = wx.App()
dialog = wx.FileDialog(None, "Choose a excel file:", style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
path = ""
if dialog.ShowModal() == wx.ID_OK:
    print("file path: ", dialog.GetPath())
    path = dialog.GetPath()
dialog.Destroy()

xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.DisplayAlerts = False  # 关闭警告
xlApp.Visible = True  # 程序可见

direct = os.path.split(path)[0]  # 获取文件夹路径 C:\PyExcel
fileName = os.path.split(path)[1]  # 获取文件名 test.xlsx
ext = os.path.splitext(path)[1]  # 获取文件拓展名 .xlsx

xlBook = xlApp.Workbooks.Open(path)  # 打开工作簿
sheetNames = [sheet.Name for sheet in xlBook.Sheets]
print("請選擇工作表", sheetNames)
mySheet = input()  # 選擇工作表
sheetDataInt = []
sheetDataFloat = []
xlBook.Worksheets(mySheet).Activate()  # 啟動工作表
sht = xlBook.Worksheets(mySheet)

LastRow = sht.usedrange.rows.count  # 有效行數
LastColumn = sht.usedrange.columns.count  # 有效列數
print("row: {}, col: {}".format(LastRow, LastColumn))

# 取出數值
for r in range(1, LastRow + 2):
    for c in range(1, LastColumn + 2):
        cell = sht.Cells(r, c)
        if cell.Interior.ColorIndex == 6:  # 如果是黃色格子
            if cell.Value % 1 == 0:
                sheetDataInt.append(int(cell.Value))
            else:
                sheetDataFloat.append(cell.Value)
print("int:", sheetDataInt)
print("float:", sheetDataFloat)
