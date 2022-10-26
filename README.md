# bot-ekin-sibon

Bot ini berfungsi untuk melakukan input data kedalam website "https://ekinerja.situbondokab.go.id/"
Penginputan dilakukan dengan membaca isi dari file excel, kemudian akan mengisikan data tersebut pada form input tambah aktivitas harian

*disarankan untutk menggunakan browser google chrome sebagai default browser anda saat menggunakan bi ini.

Hal-hal yang perlu dilakukan sebelum menggunakan bot ini :
1. install python pada komputer anda
2. install library selenium untuk mengakses browser
3. install library openpyxl untuk mengelola file excel
4. siapkan file excel yang berisi data yang akan diinputkan kedalam website e-kin
   nantinya file ini direncakakan akan dibuat secara otomatis melalui website "Chitesh", sehingga dapat memudahkan penyiapan file yang akan digunakan oleh bot ini.
   
Cara menggunakan bot e-kin :
1. simpan file excel kedalam folder yang sama dengan file bot-ekin.py
2. jalankan bot dengan klik 2 kali pada file bot-ekin-dinamis.py
3. jika ingin melakukan input ekin, pilih menu nomor 1, dengan menginputkan angka 1 dan tekan enter
4. jika ingin melakukan pengesahan ekin, pilih menu nomor 2, dengan menginputkan angka 2 dan tekan enter,
   kemudian memasukkan username dan password e-kinerja(ekin) anda.
5. bot akan mengerjakan tugasnya, silahkan tunggu hingga selesai.
