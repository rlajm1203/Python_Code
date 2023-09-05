from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import time
from getpass import getpass

"""def WebDriverSetting():
  download_path = r'C:\Downloads'
  options = webdriver.ChromeOptions()
  options.add_experimental_option('prefs', {
      'download.default_directory': download_path,
      'download.prompt_for_download': False,
      'download.directory_upgrade': True,
      'safebrowsing.enabled': True
  })
  driver = webdriver.Chrome(options=options)

  return driver

def login(ID, Password, driver=WebDriverSetting()):
  driver.get("https://www.kbid.co.kr/")
  time.sleep(1)
  driver.find_element(By.XPATH, "//*[@id='userid']").send_keys(kbid_id)
  driver.find_element(By.XPATH, "//*[@id='password']").send_keys(kbid_pw)
  driver.find_element(By.XPATH, "//*[@id='m-container']/div[3]/form/button").click()
  #driver.find_element(By.XPATH, "//*[@id='m-container']/div[3]/form/button").submit()

  return driver



def Download(driver):
  for i in range(1,10000):
    
    print(str(i)+"번째 페이지 다운로드중..")

    # 웹 페이지 접속
    url = f"https://www.kbid.co.kr/common/main_search_result.htm?CurSelPage={i}&lstFindList=1&Desc=desc&GetUp=&GetTname=C&GetArea=&GetSArea=&Kind_type=2&FindG2b=&FindG2bFull=&Tname=C&lstWork=&lstKind=&lstArea={999999}&lstSArea=&rdoFindDate=1&txtSDate={start_date}&txtEDate={end_date}&rdoFindWord=1&txtFindWord=&gCode=&lstViewList=100"
    driver.get(url)

    # 쿠키 추가
    #for name, value in cookie.items():
        #driver.add_cookie({'name': name, 'value': value})

    #체크박스 체크
    checkbox = driver.find_elements(By.XPATH, "//*[@id='allCheck']")
    checkbox.click()

    # 파일 다운로드 버튼 클릭
    #driver.switch_to.CLASS_NAME('list-function')
    button = driver.find_element(By.XPATH, "//input[@value='선택엑셀출력']")
    button.click()

    #파일 이름
    file_name = "KBID_결과_" + time.strftime("%Y%m%d%H%M%S") + ".xls"

    # 파일 다운로드 완료 대기
    file_path = os.path.join(download_path, file_name)
    while True:
      if os.path.exists(file_path):
          break
      time.sleep(1)"""

# 사용자로부터 아이디와 비밀번호 입력 받음
kbid_id = input("KBID ID를 입력하세요: ")
kbid_pw = getpass("KBID PW를 입력하세요: ")

#날짜 범위 입력
start_date, end_date = map(str,input("시작 날짜, 끝 날짜를 입력(2021-01-01 2021-12-31) : ").split())

#Download(login(kbid_id, kbid_pw))

"""cookie = {
    "KBID_ID" : kbid_id,
    "KBID_PW" : kbid_pw,
    "PHPSESSID" : "4efs9huqkqqg7veb3o9e20uic4"
}"""



# 웹 드라이버 설정
download_path = r'C:\Users\user\Downloads'
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': download_path,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})
driver = webdriver.Chrome(options=options)



# 로그인
driver.get("https://www.kbid.co.kr/")
time.sleep(1)
driver.find_element(By.XPATH, "//*[@id='userid']").send_keys(kbid_id)
driver.find_element(By.XPATH, "//*[@id='password']").send_keys(kbid_pw)
driver.find_element(By.XPATH, "//*[@id='m-container']/div[3]/form/button").click()
#driver.find_element(By.XPATH, "//*[@id='m-container']/div[3]/form/button").submit()

for i in range(1,10000):
  
  print(str(i)+"번째 페이지 다운로드중..")

  # 웹 페이지 접속
  url = f"https://www.kbid.co.kr/common/main_search_result.htm?CurSelPage={i}&lstFindList=1&Desc=desc&GetUp=&GetTname=C&GetArea=&GetSArea=&Kind_type=2&FindG2b=&FindG2bFull=&Tname=C&lstWork=&lstKind=&lstArea={999999}&lstSArea=&rdoFindDate=1&txtSDate={start_date}&txtEDate={end_date}&rdoFindWord=1&txtFindWord=&gCode=&lstViewList=100"
  driver.get(url)

  # 쿠키 추가
  #for name, value in cookie.items():
      #driver.add_cookie({'name': name, 'value': value})

  #체크박스 체크
  checkboxes = driver.find_elements(By.XPATH, "//*[@id='allCheck']")
  for checkbox in checkboxes:
    checkbox.click()

  # 파일 다운로드 버튼 클릭
  #driver.switch_to.CLASS_NAME('list-function')
  button = driver.find_element(By.XPATH, "//input[@value='선택엑셀출력']")
  button.click()

  #파일 이름
  file_name = "KBID_결과_" + time.strftime("%Y%m%d") + str(i) +"번째 페이지" + ".xls"

  # 파일 다운로드 완료 대기
  """file_path = os.path.join(download_path, file_name)
  while True:
    if os.path.exists(file_path):
        break
    time.sleep(1)"""

#로그아웃
logout = driver.find_elements(By.XPATH, "//*[@id='header'']/div[1]/div/ul/li[2]/a")
logout.click()