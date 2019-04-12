import win32api
import time

dif = 1
exc = False

while True:
    try:
        pos = win32api.GetCursorPos()
        print (pos)
        time.sleep(10)
        curPos = win32api.GetCursorPos()
        
        if pos == curPos:
            print ('mouse did not move')
            newPos = (curPos[0], curPos[1] + 10)
            dif *= -1
            print('new postion is ')
            print(newPos)
            win32api.SetCursorPos((455,456))
            print('wowow')
            exc = False
    except Exception as e:
        #if not exc:
            print('Damn error!' + str(e))
            exc = True