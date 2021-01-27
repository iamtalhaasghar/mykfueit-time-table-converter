import pyautogui as pg
import os
import datetime
import time
start = datetime.datetime.now()
t = 0.25
os.popen('test.xls')
time.sleep(2)
pg.press(['left','enter'])
pg.moveTo(22,34)    # File menu
pg.click(interval=t)
pg.moveTo(45,235)   # save as
pg.click(interval=t)
pg.moveTo(496,210)  # where: current folder
pg.click(interval=t)
pg.typewrite('myfilename')
pg.moveTo(305, 397)  # select: save as type
pg.click(interval=t)
pg.moveTo(305, 414)  # Excel Workbook
pg.click(interval=t)
pg.hotkey('alt','s')
pg.hotkey('alt','F4')
end = datetime.datetime.now()
print((end-start).total_seconds())
input()
