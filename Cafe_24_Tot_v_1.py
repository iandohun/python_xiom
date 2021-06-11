# 3팀 전용 OMS ID : online / PW : online1908

import time
import pyautogui
import pyperclip
import pandas as pd
import os
import datetime
import xlrd, xlwt
from prettytable import PrettyTable
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import sys
from selenium import webdriver
import selenium
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common import action_chains
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
ti1 = time.time()
driver = webdriver.Chrome("./chromedriver",options=options)
driver.implicitly_wait(10)
action = ActionChains(driver)
timeeorl = 1
options = Options()


# 카페 24 다운 파일
cafe24_file = 'C:/kwakcode/01_oms_to_cafe24_worh/oms_inventory/카페24_원본/cafe24_210610.xlsx'
cafe24_df = pd.read_excel(cafe24_file) # 카페24 프레임 워크


date_map_file = 'C:/kwakcode/data_map/date_map_code_v_1_8_7.xlsx'
data_df = pd.read_excel(date_map_file) # 데이터 맵 프레임 워크

def timeset():
    driver.implicitly_wait(5)
    time.sleep(3)    

url = "http://oms.xiomtt.com/"
driver.get(url)
driver.maximize_window ()
driver.find_element_by_id("f_id-inputEl").send_keys("online",Keys.TAB,"online1908")
timeset()
driver.find_element_by_id("image-1009").click()
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[3]/div/table/tbody/tr[4]/td/div/span").click()
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[3]/div/table/tbody/tr[7]/td/div/span").click() #재고 할당관리 클릭
timeset()
driver.switch_to_default_content()
driver.switch_to_frame(0)
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/span/div/table[1]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").click()
timeset()
driver.find_element_by_xpath("/html/body/div[4]/div/ul/li[2]").click() # 화주사 XIOM
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/span/div/table[5]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").click() # 품목 그룹 코드 클릭
timeset()
driver.find_element_by_xpath("/html/body/div[7]/div/ul/li[1]").click() # RUBBER 클릭
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/span/div/table[6]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").click() # 출고유형
timeset()
driver.find_elements_by_class_name("x-boundlist-item")[11].click()
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/span/div/table[8]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").clear()
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div[11]/div/a/span[1]").click() # 조회버튼
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").send_keys("10000",Keys.ENTER)
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div[11]/div/a/span[1]").click() # 조회버튼
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div[13]/div/a/span[1]").click() # 엑셀 저장 버튼
def johun(i):
    driver.find_element_by_xpath("/html/body/div[2]/div[2]/span/div/table[5]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").click() # 품목 그룹 코드 클릭
    timeset()
    driver.find_element_by_xpath("/html/body/div[7]/div/ul/li[" + str(i) + "]").click() # BLADE 클릭
    timeset()
    driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div[11]/div/a/span[1]").click() # 조회버튼
    timeset()
    driver.find_element_by_xpath("/html/body/div[1]/div/div/div[1]/div/div/div[13]/div/a/span[1]").click() # 엑셀 저장 버튼
    timeset()

for i in range(2,8,1) :    
    johun(i)
download_fold = 'C:/Users/iankw/Downloads'

xiom_omsfile_hoabo_1 = download_fold + '/재고정보관리-재고할당관리.xls'
xiom_df_hoabo_1 = pd.read_excel(xiom_omsfile_hoabo_1)
timeset()
xiom_omsfile_hoabo_2 = download_fold + '/재고정보관리-재고할당관리 (1).xls'
xiom_df_hoabo_2 = pd.read_excel(xiom_omsfile_hoabo_2)
timeset()
xiom_omsfile_hoabo_3 = download_fold + '/재고정보관리-재고할당관리 (2).xls'
xiom_df_hoabo_3 = pd.read_excel(xiom_omsfile_hoabo_3)
timeset()
xiom_omsfile_hoabo_4 = download_fold + '/재고정보관리-재고할당관리 (3).xls'
xiom_df_hoabo_4 = pd.read_excel(xiom_omsfile_hoabo_4)
timeset()
xiom_omsfile_hoabo_5 = download_fold + '/재고정보관리-재고할당관리 (4).xls'
xiom_df_hoabo_5 = pd.read_excel(xiom_omsfile_hoabo_5)
timeset()
xiom_omsfile_hoabo_6 = download_fold + '/재고정보관리-재고할당관리 (5).xls'
xiom_df_hoabo_6 = pd.read_excel(xiom_omsfile_hoabo_6)
timeset()
xiom_omsfile_hoabo_7 = download_fold + '/재고정보관리-재고할당관리 (6).xls'
xiom_df_hoabo_7 = pd.read_excel(xiom_omsfile_hoabo_7)
timeset()

full_oms_hoabo = [xiom_df_hoabo_1,xiom_df_hoabo_2,xiom_df_hoabo_3,xiom_df_hoabo_4,xiom_df_hoabo_5,xiom_df_hoabo_6,xiom_df_hoabo_7] # 데이터 프레임 리스트

full_oms_all_hoabo = pd.concat(full_oms_hoabo) # 데이터 프레임 하나로 만들기 / 확보 데이터 프레임 파일
print("확보된 데이터 프레임 파일 full_oms_all_hoabo - 예비할당잔여량")

os.remove(xiom_omsfile_hoabo_1)
os.remove(xiom_omsfile_hoabo_2)
os.remove(xiom_omsfile_hoabo_3)
os.remove(xiom_omsfile_hoabo_4)
os.remove(xiom_omsfile_hoabo_5)
os.remove(xiom_omsfile_hoabo_6)
os.remove(xiom_omsfile_hoabo_7)

time.sleep(1)

driver.switch_to_default_content()
driver.find_element_by_class_name("x-tab-close-btn").click()
timeset()
driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div/div/div[3]/div/table/tbody/tr[5]/td/div/span").click() # 조회 클릭
timeset()
driver.switch_to_frame(0)
driver.find_element_by_xpath("/html/body/div[4]/div[2]/span/div/table[1]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[2]").click()
timeset()
driver.find_element_by_xpath("/html/body/div[6]/div/ul/li[2]").click()
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div[11]/div/a/span[1]").click()
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[1]/input").send_keys('10000',Keys.ENTER) # 페이지당 갯수 확장 후 엔터
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div[12]/div/a").click() # 엑셀파일로 저장하기
timeset()
driver.find_element_by_xpath("/html/body/div[4]/div[2]/span/div/table[1]/tbody/tr/td[2]/div/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
timeset()
driver.find_element_by_xpath("/html/body/div[6]/div/ul/li[1]").click()
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div[11]/div/a/span[1]").click()
timeset()
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div[12]/div/a").click() # 엑셀파일로 저장하기
timeset()
xiom_omsfile = download_fold + '/재고정보관리-재고조회.xls'
xiom_df = pd.read_excel(xiom_omsfile)
timeset()
cham_omsfile = download_fold + '/재고정보관리-재고조회 (1).xls'
cham_df = pd.read_excel(cham_omsfile)
timeset()

os.remove(xiom_omsfile) # 파일 지우기
os.remove(cham_omsfile) # 파일 지우기

full_oms = [xiom_df,cham_df] # 데이터 프레임 리스트
full_oms_worh = pd.concat(full_oms) # 데이터 프레임 하나로 만들기
full_oms_worh.head() # 두번째 헤드 날리기
print("미 확보 재고 리스트 full_oms_worh - 가용수량 ")

for i in range(len(cafe24_df)):
    cafe24_df.iloc[i,5] # 카페 24 품목코드
    cafe24_df.iloc[i,9] = 0
    for j in (range(len(data_df))) :
        if cafe24_df.iloc[i,5] == data_df.iloc[j,3] :
            print('Data_Map Cafe24 상품명 : ' + str(data_df.iloc[j,0]) + ' / ' + str(data_df.iloc[j,1]))
            print('Data_Map Cafe24 품목코드 : ' + str(data_df.iloc[j,3]))
            print('Data_Map OMS 품목코드 : ' + str(data_df.iloc[j,6]))
            data_sch = data_df.iloc[j,6]
            break
    for k in (range(len(full_oms_all_hoabo))) : # 확보재고 반복문
        if data_sch == full_oms_all_hoabo.iloc[k,3]:
            print('확보된 재고 수량 : ' + str(full_oms_all_hoabo.iloc[k,16]))
            cafe24_df.iloc[i,9] = full_oms_all_hoabo.iloc[k,16]
            break
    if str(full_oms_all_hoabo.iloc[k,16]) == str('0'): # 확보재고 0 선택문
        for l in range(len(full_oms_worh)): # 비확보재고 반복문
            if data_sch == full_oms_worh.iloc[l,7]:
                print('비확보된 재고 수량 :' + str(full_oms_worh.iloc[l,14]))
                cafe24_df.iloc[i,9] = full_oms_worh.iloc[l,14]
                break

now = datetime.datetime.now()
nowDate = now.strftime('%Y-%m-%d')
cafe24_file_name = 'cafe24_' + nowDate + '_python.xlsx'
# cafe24_df.to_excel('C:/kwakcode/01_oms_to_cafe24_worh/oms_inventory/카페24_재고적용/' + cafe24_file_name, index=False)

cafe24_df_1 = cafe24_df.iloc[:900,:]
cafe24_df_2 = cafe24_df.iloc[900:,:]
cafe24_df_1.to_excel('C:/kwakcode/01_oms_to_cafe24_worh/oms_inventory/카페24_재고적용/1_' + cafe24_file_name, index=False)
cafe24_df_2.to_excel('C:/kwakcode/01_oms_to_cafe24_worh/oms_inventory/카페24_재고적용/2_' + cafe24_file_name, index=False)

