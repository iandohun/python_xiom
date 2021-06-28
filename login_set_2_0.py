
# 관리자 페이지 쉬운 로그인
# V1.0 관리자 페이지 선택 로그인
# 차후 계획 


# 관리자 페이지 쉬운 로그인
# V1.0 관리자 페이지 선택 로그인
# 차후 계획 
# v2.0 관리자 페이지 뿐 아니라 업무관련된 부분을 GUI 로 간단하게 작업할수 있게 계획

# 실행 파일 옵션 pyinstaller --noconsole --onefile c:/kwakcode/login_set/login_set_2_0.py

# 추가 계획 페이지 https://accounts.kakao.com/login/kakaobusiness?continue=https%3A%2F%2Fcenter-pf.kakao.com%2F_Dlyxaxb%2Fchats%3Fstay_signed_in%3D0 카카오 비지니스 계정




import time
from tkinter.constants import TOP
import pandas as pd
import datetime
import xlrd, xlwt
from prettytable import PrettyTable
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common import action_chains
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter


date_map_file = 'C:/kwakcode/login_set/MEGAM_OPEN_ID_20210510.xlsx'
juso_data_df = pd.read_excel(date_map_file) # 데이터 맵 프레임 워크

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])

hostname = 'Login / 20210510'
table = PrettyTable()
table.title = hostname
table.field_names = ['번호','쇼핑몰','URL (주소)','아이디','비밀번호']
for i in range(len(juso_data_df)):
    table.add_row([str(i),str(juso_data_df.iloc[i,0]),str(juso_data_df.iloc[i,2]),str(juso_data_df.iloc[i,3]),str(juso_data_df.iloc[i,4])])


print(table)


window=tkinter.Tk()

window.title("Login SM 3팀 Ian.kwak")

window.geometry("900x600+-1100+100")
window.resizable(False, False)




# 카페 24
def cafe24_login():
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = str(juso_data_df['url'][0])
    driver.get(url)
    driver.maximize_window()
    driver.find_element_by_xpath("/html/body/div[2]/div/section/div/form/div/div[1]/div/div[1]/div/input").send_keys(juso_data_df['username'][0],Keys.TAB,juso_data_df['password'][0])
    driver.find_element_by_xpath("/html/body/div[2]/div/section/div/form/div/div[3]/button").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[2]/div/div/div[1]/div[4]/a[2]").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div/div[1]/div/div[1]/div/strong[2]").click()
    try :
        driver.implicitly_wait(5)
        driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[4]/div/button").click()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/span[2]/div/div[1]/div/div[2]/img").click()
    except:
        pass


# 빠빠빠 로그인
def baba_login():
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = "https://cafe.daum.net/bbabbabbatakgu"
    driver.get(url)
    driver.maximize_window()

def getTextInput():
    result=textExample.get("1.0", tkinter.END+"-1c")
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = str(juso_data_df['url'][0])
    driver.get(url)
    driver.maximize_window()
    driver.find_element_by_xpath("/html/body/div[2]/div/section/div/form/div/div[1]/div/div[1]/div/input").send_keys(juso_data_df['username'][0],Keys.TAB,juso_data_df['password'][0])
    driver.find_element_by_xpath("/html/body/div[2]/div/section/div/form/div/div[3]/button").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[2]/div/div/div[1]/div[4]/a[2]").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div/div[1]/div/div[1]/div/strong[2]").click()
    try :
        driver.implicitly_wait(5)
        driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[4]/div/button").click()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/span[2]/div/div[1]/div/div[2]/img").click()
    except:
        pass

    driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[2]/ul/li[3]/a").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div/div[2]/ul/li[1]/a").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[1]/div[2]/div[2]/table/tbody/tr[3]/td/a[8]/span").click()
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[1]/div[2]/div[2]/table/tbody/tr[2]/td/div/div/select/option[8]").click()
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[1]/div[2]/div[2]/table/tbody/tr[2]/td/div/div/input").send_keys(result,Keys.ENTER)

def DcInsid():
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = "https://gall.dcinside.com/mgallery/board/lists?id=tabletennis"
    driver.get(url)
    driver.maximize_window()

def NaverPay():
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = "https://admin.pay.naver.com/home"
    driver.get(url)
    driver.maximize_window()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[3]/div[2]/div[2]/div[1]/ul/li[2]").click()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[2]/div[3]/div/form/fieldset/div[1]/div[1]/span/input").send_keys(juso_data_df['username'][16],Keys.TAB,juso_data_df['password'][16])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[2]/div[3]/div/form/fieldset/input").click()
    
def Cacabens():
    driver = webdriver.Chrome("c:/kwakcode/chromedriver")
    driver.implicitly_wait(10)
    action = ActionChains(driver)
    timeeorl = 1
    options = Options()
    url = "https://accounts.kakao.com/login/kakaobusiness?continue=https%3A%2F%2Fcenter-pf.kakao.com%2F_Dlyxaxb%2Fchats%3Fstay_signed_in%3D0"
    driver.get(url)
    driver.maximize_window()
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div/div/div[2]/form/fieldset/div[2]/div/input").send_keys(juso_data_df['username'][17],Keys.TAB,juso_data_df['password'][17])
    driver.implicitly_wait(5)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div/div/div[2]/form/fieldset/div[8]/button[1]").click()

page_no_1 = int(55)
page_botton_1 = int(395)

# 카페 24 관리자 페이지
label=tkinter.Label(window, text="  카페 24 관리자 페이지", width=62, height=2, fg="black", relief="groove",anchor='w')
label.place(x=page_no_1,y=40)

botton = tkinter.Button(window, width=10,height=1, text='Login', overrelief="solid", relief="groove" , command=cafe24_login) # command 함수 사용 괄호 제외
botton.place(x=page_botton_1,y=45)

# 카페 24 주문 찾기 페이지
label_1=tkinter.Label(window, text="  카페 24 주문(고객) 찾기 페이지", width=62, height=2, fg="black", relief="groove",anchor='w')
label_1.place(x=page_no_1,y=80)

textExample=tkinter.Text(window,height=1,width=15)
textExample.place(x=260,y=88)

botton_1 = tkinter.Button(window, width=10,height=1, text='찾기', overrelief="solid", relief="groove" , command=getTextInput) # command 함수 사용 괄호 제외
botton_1.place(x=page_botton_1,y=86)

# 네이버페이 관리자 페이지
labe2=tkinter.Label(window, text="  네이버 페이 관리자 페이지", width=62, height=2, fg="black", relief="groove",anchor='w')
labe2.place(x=page_no_1,y=120)

botton_2 = tkinter.Button(window, width=10,height=1, text='Login', overrelief="solid", relief="groove" , command=NaverPay) # command 함수 사용 괄호 제외
botton_2.place(x=page_botton_1,y=125)

# 카카오비지니스 관리자 페이지
labe2=tkinter.Label(window, text="  카카오 비지니스 관리자 페이지", width=62, height=2, fg="black", relief="groove",anchor='w')
labe2.place(x=page_no_1,y=160)

botton_2 = tkinter.Button(window, width=10,height=1, text='Login', overrelief="solid", relief="groove" , command=Cacabens) # command 함수 사용 괄호 제외
botton_2.place(x=page_botton_1,y=165)





# 빠빠빠 로그인 페이지
label=tkinter.Label(window, text="  빠빠빠 로그인 페이지", width=43, height=2, fg="black", relief="groove",anchor='w')
label.place(x=510,y=40)

botton = tkinter.Button(window, width=10,height=1, text='Login', overrelief="solid", relief="groove" , command=baba_login) # command 함수 사용 괄호 제외
botton.place(x=720,y=45)

# 디씨 인싸이드 로그인 페이지
label=tkinter.Label(window, text="  디씨인사이드 로그인 페이지", width=43, height=2, fg="black", relief="groove",anchor='w')
label.place(x=510,y=80)

botton = tkinter.Button(window, width=10,height=1, text='Login', overrelief="solid", relief="groove" , command=DcInsid) # command 함수 사용 괄호 제외
botton.place(x=720,y=86)







window.mainloop()
