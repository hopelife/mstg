import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('D:\\moon\\dev\\projects\\mstg\\com\\test.xlsx')
ws = wb.ActiveSheet
ws.Cells(1,1).Value = "python"
ws.Cells(1,2).Value = "is"
ws.Range("C1").Value = "good"
ws.Range("C1").Interior.ColorIndex = 10

ws.Range("A2:C2").Interior.ColorIndex = 27