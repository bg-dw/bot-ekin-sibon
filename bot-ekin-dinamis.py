# from logging import exception

from selenium import webdriver

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import time

print("Menu :")
print("1. Input E-Kin")
print("2. Verivikasi E-Kin")
pil = input("Masukkan Nomor: ")
if(pil=="1"):
    file = input("Nama File(.xls/xlsx):")
    path="D:\\Develope\\bot-ekin\\"
    wb = load_workbook(filename=str(path+file))#'r' refers to 'raw string'

    sheet = wb['Sheet1']
    row_count = len(sheet['A'])
    
    driver = webdriver.Chrome()#put your web driver into directory "C:\Windows\webdriver.exe"
    driver.get("https://ekinerja.situbondokab.go.id/")
    driver.maximize_window()
    driver.implicitly_wait(10)

    time.sleep(1)
    driver.find_element(
        'id', 'login-form__username').send_keys("xxxxx")  # silahkan ganti xxxx dengan username ekin anda
    driver.find_element(
        'id', 'login-form__password').send_keys("yyyyy")  # silahkan ganti yyyy dengan password ekin anda
    driver.find_element(
        By.XPATH, '//*[@id="login-form"]/button').click()
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-harian.html")
    driver.refresh()
    driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-harian.html")

    print(row_count)
    # # perulangan input
    i = 2

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
                ak = Select(driver.find_element(By.ID, 'aktivitasid'))
                ak.select_by_value(str(akt))
                kg = Select(driver.find_element(By.ID, 'kegiatanid'))
                kg.select_by_value(str(keg))

                driver.find_element('id', 'catatan_aktivitas').send_keys(ctt)
                driver.find_element('id', 'save-aktifitas').click()

                time.sleep(1)
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
    print("pilihan 2")
