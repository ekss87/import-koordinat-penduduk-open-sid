from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import WebDriverWait
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains

URL='https://webdesa.desa.id'

driver = webdriver.Chrome()
driver.get(URL)
driver.maximize_window()
driver.implicitly_wait(15)
driver.find_element_by_name("username").send_keys("username")
driver.find_element_by_name("password").send_keys("pasword")
driver.find_element_by_class_name("btn").click()
driver.find_element_by_link_text("Kependudukan").click()
driver.find_element_by_link_text("Penduduk").click()


now = time.strftime('%x')
#print(now)


wb = load_workbook(filename ="D://dataExcel.xlsx")
sheetRange = wb['1']

i = 2

while i <= len(sheetRange['A']):

   SA = sheetRange['A' + str(i)].value
   NIK = sheetRange['D'+ str(i)].value
   LAT = sheetRange['E'+ str(i)].value
   LNG = sheetRange['F'+ str(i)].value

   driver.find_element_by_name("cari").send_keys(NIK)
   driver.find_element_by_xpath('//*[@id="mainform"]/div[1]/div[2]/div/div/button').click()
   
   driver.find_element_by_xpath('//*[@id="mainform"]/div[2]/table/tbody/tr/td[3]/div/button').click()
   
   driver.find_element_by_xpath('//*[@id="mainform"]/div[2]/table/tbody/tr[1]/td[3]/div/ul/li[3]/a').click()
   
   driver.find_element_by_xpath('//*[@id="validasi1"]/div[2]/div/a[2]').click()
   #LOK1 = driver.find_element_by_name("lat")
   #ActionChains.double_click(LOK1)
   driver.find_element_by_xpath('//*[@id="lat"]').clear()
   driver.find_element_by_xpath('//*[@id="lat"]').send_keys(LAT)

   driver.find_element_by_id('lng').clear()
   driver.find_element_by_id('lng').send_keys(LNG)
   time.sleep(2)
   driver.find_element_by_xpath('//*[@id="simpan_penduduk"]/i').click()

   #WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, '//*[@id="maincontent"]/div[2]/div/div/div[1]/a[2]')))
   #time.sleep(2)
  
   driver.find_element_by_xpath('//*[@id="maincontent"]/div/div/div/div[1]/a[2]').click()
   sheetRange['G'+str(i)]= now
   #save dengan nama file yg sama dalam folder yang sama dengan data excel yang di import / diganti \
   wb.save(r"D:\dataExcel.xlsx")
   time.sleep(2)
   i += 1
   