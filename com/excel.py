import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Cells(1, 1).Value = "hello world"
wb.SaveAs('D:\\moon\\dev\\projects\\mstg\\com\\test.xlsx')
excel.Quit()