# excel_Add-Ins_NumberWords
Add-Ins excel yang digunakan untuk merubah nominal (angka) menjadi terbilang (kata-kata)
      ## cara import Add-Ins
      1. kita download file yang kita butuhkan terlebih dahulu di https://github.com/mydevdetex/excel_module_NumberWords/archive/refs/heads/main.zip , kemudian extract dan pastikan terdapat file number_to_words.xlam di dalam folder tersebut.
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
      ## cara penggunaan
      1. buka file excel yang diinginkan, atau bisa juga membuka file excel baru.
      2. masukkan sembarang angka pada suatu cell, misal 750000 pada cell A3.
      3. letakkan kursor di luar cell A3, misal A4.
      4. pada cell A4 tersebut kita masukkan formula dengan cara =NumberWords(A3).
      5. maka cell A4 akan menampilkan tulisan <i><b>"tujuh ratus lima puluh ribu"</i></b>.
      6. untuk membuat huruf menjadi huruf besar tiap awal kata maka tambahkan formula proper sehingga penulisan formula akan menjadi =proper(NumberWords(A3)).
      7. maka hasilnya akan menjadi <i><b>"Tujuh Ratus Lima Puluh Ribu"</i></b>.
      8. untuk menambah kata <i><b>"rupiah"</i></b> di akhir kalimat, maka formula akan menjadi =proper(NumberWords(A3)&" rupiah").
      9. maka akan menghasilkan keluaran <i><b>"Tujuh Ratus Lima Puluh Ribu Rupiah"</i></b>.

