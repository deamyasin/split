<<<<<<< HEAD
```markdown
# PDF Splitor

PDF Splitor adalah aplikasi berbasis GUI untuk memanipulasi file PDF dengan fitur seperti menandai halaman, membagi file PDF, memutar halaman, memperbesar/memperkecil tampilan, dan mengonversi file PDF ke format Word (.docx). Aplikasi ini dibangun menggunakan Python dengan pustaka seperti `tkinter`, `PyMuPDF`, dan `Pillow`.

## Fitur Utama

- **Membuka File PDF**: Muat file PDF untuk diproses.
- **Menambah Penanda Halaman**: Tandai halaman tertentu untuk membagi file PDF.
- **Menyimpan dan Memuat Penanda**: Ekspor dan impor penanda dalam format `.txt`.
- **Membagi PDF**: Bagikan file PDF berdasarkan penanda yang telah ditentukan.
- **Memutar Halaman**: Putar halaman PDF ke kiri atau ke kanan.
- **Zoom In/Out**: Perbesar atau perkecil tampilan halaman PDF.
- **Menghapus Halaman**: Hapus halaman dari dokumen PDF.
- **Konversi ke Word**: Konversi halaman PDF yang dipilih menjadi dokumen Word.
- **Navigasi Halaman Cepat**: Navigasi melalui thumbnail atau input nomor halaman.

## Persyaratan

- Python 3.9 atau lebih baru
- Pustaka berikut (dapat diinstal melalui `pip`):
  - `tkinter` (default pada Python)
  - `Pillow`
  - `PyMuPDF` (`fitz`)
  - `python-docx`

## Instalasi

1. Clone repository ini:
   ```bash
   git clone https://github.com/username/dea-pdf-splitor.git
   cd dea-pdf-splitor
   ```

2. Instal dependensi:
   ```bash
   pip install -r requirements.txt
   ```

3. Jalankan aplikasi:
   ```bash
   python main.py
   ```

## Cara Menggunakan

1. **Buka Aplikasi**: Jalankan `main.py` untuk membuka antarmuka GUI.
2. **Muat File PDF**:
   - Klik tombol `Pick File` dan pilih file PDF dari komputer Anda.
3. **Menambah Penanda**:
   - Pilih halaman pada tampilan utama.
   - Klik tombol `Add Split PDF (ENTER)` dan masukkan nama penanda.
4. **Simpan Penanda**:
   - Klik `Export Marker` untuk menyimpan penanda ke file `.txt`.
5. **Muat Penanda**:
   - Klik `Load Marker` dan pilih file penanda untuk dimuat.
6. **Membagi PDF**:
   - Klik `Proses Output`, pilih folder untuk menyimpan file hasil pembagian.
7. **Memutar Halaman**:
   - Gunakan `Shift + Left` untuk memutar halaman ke kiri.
   - Gunakan `Shift + Right` untuk memutar halaman ke kanan.
8. **Zoom In/Out**:
   - Gunakan `Shift + Plus` untuk zoom in.
   - Gunakan `Ctrl + Minus` untuk zoom out.
9. **Menghapus Halaman**:
   - Klik tombol `Delete` untuk menghapus halaman saat ini.
10. **Konversi ke Word**:
    - Klik `Convert to Word`, pilih halaman yang akan dikonversi, lalu pilih lokasi untuk menyimpan file Word.

## Navigasi Keyboard

- **Navigasi Halaman**:
  - `Left Arrow`/`Up Arrow`: Halaman sebelumnya.
  - `Right Arrow`/`Down Arrow`: Halaman berikutnya.
- **Menambah Penanda**:
  - `Enter`: Tambah penanda pada halaman saat ini.
- **Rotasi Halaman**:
  - `Shift + Left`: Putar halaman ke kiri.
  - `Shift + Right`: Putar halaman ke kanan.
- **Zoom**:
  - `Shift + Plus`: Zoom in.
  - `Ctrl + Minus`: Zoom out.

## Kontributor

Dikembangkan oleh Dea sebagai alat bantu manajemen file PDF dengan fitur-fitur unik untuk kebutuhan sehari-hari.

## Lisensi

Proyek ini dilisensikan di bawah [Lisensi MIT](LICENSE).
```
=======
# split
>>>>>>> b8ec8c99e640de37a4cc1c8e30e60f8b0e60c9e7
