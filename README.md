# DPS Excel

Pembersih karakter tak terlihat `U+200C` pada kolom NIK dan nomor KK. Semua pemrosesan berlangsung lokal.

## Setup sekali

Di macOS atau Linux, dengan Git dan Python 3.10 atau lebih baru:

```sh
git clone https://github.com/digitalpencaksilat/dps-excel.git
cd dps-excel
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

## Penggunaan

```sh
./dps-excel
./dps-excel clean "data warga.xlsx"
./dps-excel clean "data warga.xlsx" -o "hasil.xlsx"
./dps-excel --version
```

Di wizard interaktif, pilih **Tambah kolom lain?** untuk memilih kolom tambahan per sheet. Jika header NIK/KK tidak terdeteksi di sheet itu, masukkan nomor baris header. Tekan Spasi untuk menandai kolom dan Enter untuk lanjut. Kolom tambahan hanya dibersihkan dari `U+200C` dan format selnya dipertahankan. Mode `clean` tetap hanya memproses NIK/KK.

Wizard membuka pemilih file `.xlsx` bawaan macOS (`osascript`) atau Linux (`zenity`/`kdialog`). Jika dialog tidak tersedia, masukkan path di terminal. Membatalkan dialog menghentikan wizard. Setelah memilih file, wizard menampilkan detail file serta sheet dan kolom yang terdeteksi, lalu meminta konfirmasi.

Untuk menjalankan `dps-excel` dari folder mana pun, jalankan ini dari root proyek jika `~/.local/bin` sudah ada di `PATH`:

```sh
mkdir -p "$HOME/.local/bin"
ln -s "$(pwd)/dps-excel" "$HOME/.local/bin/dps-excel"
```

Launcher membaca source langsung. Tidak perlu reinstall saat kode berubah. Path relatif mengacu pada folder terminal saat perintah dijalankan.

Deteksi otomatis memakai baris pertama yang memuat header dikenal: `NIK`, `No NIK`, `No. NIK`, `No KK`, `No. KK`, `Nomor KK`, atau `Nomor Kartu Keluarga`. Perbedaan kapitalisasi dan spasi diabaikan. Pembersihan dimulai setelah baris header. Sheet tanpa kecocokan dilewati kecuali kolomnya dipilih manual. Asumsinya satu baris header per sheet.

Hasil default `<nama>_bersih.xlsx`. File yang sudah ada tidak ditimpa. Hanya sel teks biasa di kolom target yang dibersihkan; angka dan rumus dilewati. Presisi angka yang sudah hilang di Excel tidak dapat dipulihkan.

## Pengembangan

Versi mengikuti [Semantic Versioning](https://semver.org/). Saat merilis, perbarui `--version` di `src/cli.py` dan `CHANGELOG.md`, jalankan tes, lalu buat tag `v<versi>`.

```sh
.venv/bin/python -m unittest discover -s tests -v
```
