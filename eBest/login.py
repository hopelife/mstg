import win32com.client
import pythoncom

class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("로그인 성공")
            XASessionEventHandler.login_state = 1
        else:
            print("로그인 실패")

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)


f = open("D:\moon\dev\projects\_accounts\ebest.txt", "rt")
lines = f.readlines()
print(lines)

id = lines[0].rstrip('\n')
passwd = lines[1].rstrip('\n')
cert_passwd = lines[2].rstrip('\n')

print(id)

# instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
# instXASession.Login(id, passwd, cert_passwd, 0, 0)

# while XASessionEventHandler.login_state == 0:
#     pythoncom.PumpWaitingMessages()