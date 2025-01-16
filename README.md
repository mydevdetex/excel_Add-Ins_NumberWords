# excel_Add-Ins_NumberWords
Add-Ins excel yang digunakan untuk merubah nominal (angka) menjadi terbilang (kata-kata)
## cara import Add-Ins
1. kita download file yang kita butuhkan terlebih dahulu di https://github.com/mydevdetex/excel_Add-Ins_NumberWords/archive/refs/heads/main.zip , kemudian extract dan pastikan terdapat file number_to_words.xlam di dalam folder tersebut.
2. buka file excel yang diinginkan, atau bisa juga membuka file excel baru. klik menu Developer di menu bar pada file excel yang sedang dibuka. Jika menu Developer belum aktif, maka aktifkan terlebih dahulu dengan cara :
   ### aktifasi menu Developer
   a. di menu File tab, klik Options > Customize Ribbon. <br>
   b. pada Customize the Ribbon dan di bawah Main Tabs (berada di sisi kanan), pilih Developer check dan centang pada box tersebut, maka menu Developer akan tampil.
3. pilih menu Add-Ins.
4. klik Browse...
5. cari dan pilih file number_to_words.xlam, setelah itu klik Ok.
6. maka akan tampil Add-Ins baru bernama Number_To_Words
7. pastikan Add-Ins ini telah tercentang agar Add-Ins ini aktif.
8. Add-Ins pun siap digunakan.
9. Jika add-ins tidak berjalan secara otomatis setelah Excel di tutup dan di buka kembali(biasanya di versi Excel 2013 ke atas), maka lakukan langkah-langkah berikut:
   ### cara menjalankan Add-Ins secara otomatis
   a. Buka Windows File Explorer dan masuk ke folder tempat file number_to_words.xlam disimpan<br>
   b. Klik kanan file number_to_words.xlam dan pilih Properti<br>
   c. Di bagian bawah tab General (Umum), pilih kotak centang Unblock (Buka blokir)<br>
   d. Klik OK<br>
   e. Buka Excel<br>
   f. Klik menu Developer, lalu klik Add_Ins<br>
   g. Hapus centang salah satu Add-Ins, lalu klik OK<br>
   h. Mulai ulang Excel
## cara penggunaan
1. buka file excel yang diinginkan, atau bisa juga membuka file excel baru.
2. masukkan sembarang angka pada suatu cell, misal 750000 pada cell A3.
3. letakkan kursor di luar cell A3, misal A4.
4. pada cell A4 tersebut kita masukkan formula dengan cara <b>=NumberWords(A3)</b>.
5. maka cell A4 akan menampilkan tulisan <i><b>"tujuh ratus lima puluh ribu"</i></b>.
6. untuk membuat huruf menjadi huruf besar tiap awal kata maka tambahkan formula proper sehingga penulisan formula akan menjadi <b>=proper(NumberWords(A3))</b>.
7. maka hasilnya akan menjadi <i><b>"Tujuh Ratus Lima Puluh Ribu"</i></b>.
8. untuk menambah kata <i><b>"rupiah"</i></b> di akhir kalimat, maka formula akan menjadi <b>=proper(NumberWords(A3)&" rupiah")</b>.
9. maka akan menghasilkan keluaran <i><b>"Tujuh Ratus Lima Puluh Ribu Rupiah"</i></b>.
## link video tutorial
https://youtu.be/0BY31zl9LE0?si=G4t1EJDeKql1IftE