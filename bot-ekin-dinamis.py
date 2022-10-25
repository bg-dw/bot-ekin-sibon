# from logging import exception

from ast import Continue
from selenium import webdriver

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import os 
import glob
import base64
import time

print("Menu :")
print("1. Input E-Kin")
print("2. Verifikasi E-Kin")
pil = input("Masukkan Nomor: ")

if(pil=="1"):
    dir_path = os.path.dirname(os.path.realpath(__file__))#path untuk file bot
    list_of_files = glob.glob(dir_path+"\*.xlsx") # * artinya semua file, *.xlsx untuk semua file dengan eksistensi tersebut
    latest_file = max(list_of_files, key=os.path.getctime)#mendapatkan file terakhir
    wb = load_workbook(filename=str(latest_file))#'r' refers to 'raw string'

    sheet = wb['Sheet1']#sheet terpilih
    row_count = len(sheet['A'])#menghitung panjang baris
    
    driver = webdriver.Chrome()#put your web driver into directory "C:\Windows\webdriver.exe"
    driver.get("https://ekinerja.situbondokab.go.id/")
    driver.maximize_window()
    driver.implicitly_wait(10)
    
    #mengambil username dan password
    uname = sheet['A1'].value
    pw = sheet['B1'].value
    
    time.sleep(1)
    driver.find_element(
        'id', 'login-form__username').send_keys(uname)  # silahkan ganti xxxx dengan username ekin anda
    time.sleep(1)
    driver.find_element(
        'id', 'login-form__password').send_keys(pw)  # silahkan ganti yyyy dengan password ekin anda
    time.sleep(1)
    driver.find_element(
        By.XPATH, '//*[@id="login-form"]/button').click()
    
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-harian.html")
    driver.refresh()
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-harian.html")#mengantisipasi kegagalan load halaman
    

    #perulangan input
    i = 3

    while i <= row_count:
        tgl = sheet['A'+str(i)].value
        akt = sheet['B'+str(i)].value
        keg = sheet['C'+str(i)].value
        ctt = sheet['D'+str(i)].value

        try:
            if ((i-1) != row_count):
                driver.find_element(
                    By.XPATH, '//*[@id="main-content"]/div[1]/div[2]/article/div/div/div[2]/button').click()
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located(
                    (By.XPATH, '//*[@id="aktivitas-harian-modal"]/div/div')))

                driver.find_element('id', 'tgl_aktivitas').send_keys(tgl)
                time.sleep(1)
                ak = Select(driver.find_element(By.ID, 'aktivitasid'))
                ak.select_by_value(str(akt))
                time.sleep(1)
                kg = Select(driver.find_element(By.ID, 'kegiatanid'))
                kg.select_by_value(str(keg))
                time.sleep(1)
                driver.find_element('id', 'catatan_aktivitas').send_keys(ctt)
                time.sleep(1)
                driver.find_element('id', 'save-aktifitas').click()
                driver.refresh()
            else:
                break
        except TimeoutException:
            print("Gagal")

        i += 1

        x = i-2
        time.sleep(2)
        print("Input data ke-"+(str(x))+" selesai!")
        if (x+1 == row_count):
            print("Semua data berhasil disimpan!")

else:
    user = input("Masukkan Username: ")
    pas = input("Masukkan Password: ")
    
    driver = webdriver.Chrome()#put your web driver into directory "C:\Windows\webdriver.exe"
    driver.get("https://ekinerja.situbondokab.go.id/")
    driver.maximize_window()
    driver.implicitly_wait(10)

    time.sleep(1)
    driver.find_element(
        'id', 'login-form__username').send_keys(user)  # silahkan ganti xxxx dengan username ekin anda
    time.sleep(1)
    driver.find_element(
        'id', 'login-form__password').send_keys(pas)  # silahkan ganti yyyy dengan password ekin anda
    time.sleep(1)
    driver.find_element(
        By.XPATH, '//*[@id="login-form"]/button').click()
    
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-pengesahan.html")
    driver.refresh()
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-pengesahan.html")#mengantisipasi kegagalan load halaman
    
    list_hal = driver.find_elements(By.XPATH, '//*[@id="list-table"]/article/div/div/div[3]/div/ul')
    if(list_hal):#jika element halaman ditemukan
        halaman = len(list_hal[0].find_elements(By.TAG_NAME, 'li'))
    else:
        halaman=1
    
    x=1
    while x<=halaman:
        list_bar = driver.find_elements(By.XPATH, '//*[@id="list-table"]/article/div/div/div[1]/table/tbody')
        if(list_bar):#jika tabel ditemukan
            bar = len(list_bar[0].find_elements(By.TAG_NAME, 'tr'))
        else:
            bar=1
        tr = 1
        while tr<=bar:
            try:
                if(driver.find_element(By.XPATH, '//*[@id="list-table"]/article/div/div/div[1]/table/tbody/tr[1]/td[6]')):
                    id_form = driver.find_element(
                        By.XPATH, '//*[@id="list-table"]/article/div/div/div[1]/table/tbody/tr[1]/td[6]').get_attribute("data-target").split("-")
                    driver.find_element(
                        By.XPATH, '//*[@id="list-table"]/article/div/div/div[1]/table/tbody/tr[1]/td[6]').click()
                    time.sleep(1)
                    driver.find_element(
                        By.XPATH, '//*[@id="form-pengesahan-'+id_form[-1]+'"]/div[5]/div/a[2]').click()
                    driver.refresh()
            except NoSuchElementException:
                print("Tugas Selesai!")  
            tr+=1 
        x+=1
