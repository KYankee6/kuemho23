'''
Author : 양준혁
Author-email : didzl1231@naver.com
Version : 1.03
Build Tools : Python 3.7
Date it was made: 2020-01-08
last update : 2020-01-16
'''

import win32com.client
import datetime
import time
import os
from selenium import webdriver
import pyautogui
import sys
excel = win32com.client.Dispatch("Excel.Application")

def printer(path):
time.sleep(1)
wb = excel.Workbooks.Open(path)
ws = wb.ActiveSheet
excel.Visible = True
time.sleep(1.5)
# excel_center = pyautogui.locateCenterOnScreen('c:\Users\T1\Desktop\자동정산기\excel_image.PNG')
# pyautogui.click(excel_center)
pyautogui.hotkey('winleft','d')
pyautogui.hotkey('winleft','1')
pyautogui.hotkey('alt','p','s','p')
pyautogui.press('tab')
pyautogui.press('l')
pyautogui.press('tab', 2)
pyautogui.press('6')
pyautogui.press('0')
pyautogui.press('enter')
pyautogui.hotkey('ctrl','p')
pyautogui.press('enter')
time.sleep(2)
excel.Quit()
pyautogui.press('n')

def get_jubsu_excel():
# webdriver로 엑셀 파일 다운받기
browser = webdriver.Chrome(r'C:\Users\T1\PycharmProjects\Contact\chromedriver.exe')
browser.get('http://sd.go.kr/hcms/sdadmin.do')
pyautogui.press('enter')
id = ''
pw = ''
element = browser.find_element_by_name("id")
element.send_keys(id)
element = browser.find_element_by_name("pwd")
element.send_keys(pw)
pyautogui.press('enter')
browser.get('http://sd.go.kr/hcms/waste.do')
element = browser.find_element_by_name("searchDayExcel")
element.send_keys(nowDate)
browser.find_element_by_xpath(
'/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/a').click()
time.sleep(1)
element.clear()
browser.close()
# excel.Visible = True

def get_baechul_excel():
# webdriver로 엑셀 파일 다운받기
browser = webdriver.Chrome(r'C:\Users\T1\PycharmProjects\Contact\chromedriver.exe')
browser.get('http://sd.go.kr/hcms/sdadmin.do')
pyautogui.press('enter')
id = ''
pw = ''
element = browser.find_element_by_name("id")
element.send_keys(id)
element = browser.find_element_by_name("pwd")
element.send_keys(pw)
pyautogui.press('enter')
browser.get('http://sd.go.kr/hcms/waste.do')
element = browser.find_element_by_name("searchEjectionDayExcel")
element.send_keys(nowDate)
browser.find_element_by_xpath(
'/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[6]/td[2]/a').click()
time.sleep(1)
element.clear()
browser.close()

sd.go.kr으로부터 엑셀파일을 받아옵니다
og_path = 'C:\Users\T1\Downloads\waste_'
tDate = datetime.datetime.now()
nowDate = tDate.strftime('%Y%m%d')
path = og_path + nowDate

print(path)

예외처리, 다운받기 전에 존재한다면 삭제시킵니다.
if os.path.isfile(path+'.xls'):
os.remove(path + '.xls')
print("기존 폐기물 정산 엑셀 삭제함.\n")

get_jubsu_excel()

엑셀로부터 정산시작
wb= excel.Workbooks.Open(path+'.xls')
ws=wb.ActiveSheet
excel.Visible=True
CashSum=[]
CreditSum=[]
Sum=[]
total_Cash=0
total_Credit=0
total=0
for i in range(300):
if(ws.Cells(5+i,11).Value==None):break
if(ws.Cells(5+i,3).Value=='동접수') : Sum.append(ws.Cells(5+i,11).Value)
if(ws.Cells(5+i,13).Value=='현금' and ws.Cells(5+i,3).Value=='동접수'):
CashSum.append(ws.Cells(5+i,11).Value)
elif(ws.Cells(5+i,13).Value=='카드' and ws.Cells(5+i,3).Value=='동접수'):
CreditSum.append((ws.Cells(5+i,11).Value))
excel.Quit()

f=open('c:\Users\T1\Desktop\자동정산기\일별_정산_결과\'+nowDate+'.txt','wt',encoding='UTF8')
for row in CashSum:
total_Cash+=int(row)
for row in CreditSum:
total_Credit+=int(row)
for row in Sum:
total+=int(row)

f.write("현금 총 금액: %s\n" % total_Cash)
f.write("카드 총 금액: %s\n" % total_Credit)
f.write("총 금액: %s\n" % total)
os.system('cls')
print("현금 총 금액: %s\n" % total_Cash)
print("카드 총 금액: %s\n" % total_Credit)
print("총 금액: %s\n" % total)
time.sleep(2)
doyouwannadevelopanapp=input("Do you want to print out? [y/n] : ")
if(doyouwannadevelopanapp=='y'):
os.system('cls')
print("현금 총 금액: %s\n" % total_Cash)
print("카드 총 금액: %s\n" % total_Credit)
print("총 금액: %s\n" % total)
print()
print("DO NOT INTERRUPT")
print()
print("실행 중... 키보드,마우스 쓰지 마세요.")
printer(path)
os.remove(path + '.xls')
get_baechul_excel()
printer(path)
else:
print("No print needed")
os.remove(path+'.xls')
f.close()

sys.exit()
