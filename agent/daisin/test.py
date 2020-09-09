#import win32com.client
#instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
#print(instCpCybos.IsConnect)


#import win32com.client
#instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
#print(instCpStockCode.GetCount())

#import win32com.client
#instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
#print(instCpStockCode.GetData(1, 0))


import win32com.client
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
for i in range(0, 10):
    print(instCpStockCode.GetData(1,i))