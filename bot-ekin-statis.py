# from logging import exception

# library selenium untuk komunikasi dengan browser
from selenium import webdriver

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

# library openpyxl untuk mengelola file excel
from openpyxl import load_workbook
import time

# file yang akan dibaca oleh python
wb = load_workbook(filename="C:\laragon\www\python\ekin.xlsx")

sheet = wb['Sheet1']  # sheet yang dibaca
row_count = len(sheet['A'])  # jumlah baris pada sheet excel

# put your web driver into directory "C:\Windows\webdriver.exe"
driver = webdriver.Chrome()

# url tujuan untuk input data
driver.get("https://ekinerja.situbondokab.go.id/")
driver.maximize_window()
driver.implicitly_wait(10)

time.sleep(1)  # delay sebelum eksekusi perintah selanjutnya
driver.find_element(
    'id', 'login-form__username').send_keys("xxxx")  # silahkan ganti xxxx dengan username ekin anda
time.sleep(1)
driver.find_element(
    'id', 'login-form__password').send_keys("yyyy")  # silahkan ganti yyyy dengan password ekin anda
time.sleep(1)
driver.find_element(
    By.XPATH, '//*[@id="login-form"]/button').click()

# alamat tujuan setelah login
driver.get("https://ekinerja.situbondokab.go.id/page/aktivitas-harian.html")

# perulangan input
i = 2

while i <= row_count:  # mengulang sebanyak jumlah baris pada excel worksheet
    tgl = sheet['A'+str(i)].value
    akt = sheet['B'+str(i)].value
    keg = sheet['C'+str(i)].value
    ctt = sheet['D'+str(i)].value

    try:
        if ((i-1) != row_count):  # jika belum mencapat baris terakhir
            driver.find_element(
                By.XPATH, '//*[@id="main-content"]/div[1]/div[2]/article/div/div/div[2]/button').click()  # klik button untuk menampilkan modal tambah data

            WebDriverWait(driver, 15).until(EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="aktivitas-harian-modal"]/div/div')))  # beralih pada modal

            # mengirim data excel ke form input
            driver.find_element('id', 'tgl_aktivitas').send_keys(tgl)
            ak = Select(driver.find_element(By.ID, 'aktivitasid'))
            ak.select_by_value(str(akt))
            kg = Select(driver.find_element(By.ID, 'kegiatanid'))
            kg.select_by_value(str(keg))

#             driver.find_element('id', 'catatan_aktivitas').send_keys(ctt)
#             driver.find_element('id', 'save-aktifitas').click()

            time.sleep(1)
            driver.refresh()  # refresh halaman setelah input data
        else:
            break  # menghentikan program
    except TimeoutException:  # jika terjadi kegagalan
        print("Gagal")

    i += 1
    x = i-2
    time.sleep(2)
    # mengetahui banyak data yang telah diinputkan
    print("Input data ke-"+(str(x))+" selesai!")
    if (x+1 == row_count):
        print("Semua data berhasil disimpan!")
