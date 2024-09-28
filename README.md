# PPh 21 TER Counter Ver 1.0 (2024)

![demo](https://github.com/user-attachments/assets/12d10f78-eaa6-47a9-857b-a70349646b2c)
## Overview
Modul VBA ini dirancang untuk memformat data **Pajak Penghasilan Karyawan Tetap maupun Tidak Tetap yang dibayar bulanan** sesuai dengan ketentuan **TER (Tarif Efektif Rerata) Bulanan** yang diatur dalam *Peraturan Menteri Keuangan (PMK) No. 168 Tahun 2023* tentang Tentang Petunjuk Pelaksanaan Pemotongan Pajak Atas Penghasilan Sehubungan dengan Pekerjaan, Jasa, Atau Kegiatan Orang Pribadi.
## Tujuan
Modul dapat digunakan untuk memvalidasi perhitungan PPh 21 TER secara manual di Excel sekaligus pembanding dengan perhitungan yang dilakukan secara otomatis di DJP Online.
## Fitur
- **Pemformatan Data dalam Jumlah Banyak**  
Modul dapat menentukan dan menghitung data gaji karyawan dalam jumlah yang banyak dalam beberapa detik. 
- **Dukungan File `Data TER`**  
Data set dasar pengenaan tarif TER Bulanan berdasarkan 3 kategori, yakni A, B, dan C telah tersedia dalam repositori ini.
- **Pemilihan Kolom Gaji secara Fleksibel**  
Modul akan meminta anda untuk memasukkan kolom yang berisi data gaji, dan akan membatalkan operasi jika input tidak valid.
## Persyaratan Aplikasi
- Microsoft Excel
- VBA for Excel
## Instalasi Modul
1. Buka Workbook Excel anda yang sudah dibersihkan dimana letak PTKP dan Total Penerimaan Bruto telah tersedia pada masing-masing kolom.
2. Tekan `Alt + F11` untuk membuka VBA Editor
3. Buat modul baru dengan cara klik kanan pada modul yang sudah ada atau nama workbook di Project Explorer, lalu pilih Insert > Module.
4. Salin dan tempelkan seluruh isi dari `PPhTERFunc.bas` ke dalam modul baru tersebut.
5. Simpan workbook sebagai workbook dengan macro yang diaktifkan/ *Excel Macro Enabled* (.xlsm) (opsional)
## Ketentuan PTKP<sup>1</sup> 
| Kategori  TER | PTKP                  |
|:-------------:|:---------------------:|
| A             | TK/0, TK/1, K/1       |
| B             | TK/2, TK/3, K/1, K/2  |
| C             | K/3                   |

Contoh format Penulisan yang benar disisipkan garis miring :  
> ✅ K/1  
❌ K1

<sub> 1 : PP No. 58 Tahun 2023 </sub>
## Penggunaan
1. Buka workbook Excel yang berisi data gaji karyawan yang sudah anda bersihkan.
2. Tekan `Alt + F8` untuk membuka kotak dialog Macro.
3. Pilih `formatting_PPh21TER` dari daftar makro dan klik Jalankan (Run).
4. Konfirmasi prompt untuk melanjutkan proses perhitungan.
5. Ikuti alur petunjuk yang diberikan hingga modul memunculkan hasil akhir perhitungan PPh 21.
## Catatan
⚠️ *Pastikan kolom PTKP berada tepat di sisi kiri kolom penerimaan bruto untuk menghindari kesalahan pemrosesan.*  
⚠️ *Apabila terjadi ketidaksesuaian antara hasil perhitungan modul dengan web DJP Online, periksa secara manual pada sheet `DATA TER` yang disediakan.*