# DPS Excel

**Bersihkan karakter tak terlihat dari file Excel, langsung di komputer Anda.**

[![Versi](https://img.shields.io/github/v/tag/digitalpencaksilat/dps-excel?sort=semver&label=versi)](https://github.com/digitalpencaksilat/dps-excel/tags)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)](https://www.python.org/)
![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Linux-555)

## Fitur Utama

- **Deteksi otomatis:** Temukan kolom NIK dan nomor KK di setiap sheet, meski header tidak berada di baris pertama.
- **Kolom tambahan:** Pilih kolom seperti Nama atau Alamat lewat wizard, tanpa mengubah format selnya.
- **Lokal dan tanpa menimpa:** File tidak diunggah. Hasil disimpan terpisah tanpa menimpa file yang sudah ada.

## Mulai Cepat

Di macOS atau Linux, dengan Git dan Python 3.10 atau lebih baru:

```sh
git clone https://github.com/digitalpencaksilat/dps-excel.git
cd dps-excel
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
./dps-excel
```

## Wizard Interaktif

Jalankan `./dps-excel` lalu pilih file `.xlsx`. Wizard membuka dialog bawaan macOS (`osascript`) atau Linux (`zenity`/`kdialog`); jika dialog tidak tersedia, masukkan path di terminal. Membatalkan dialog menghentikan wizard.

Untuk memproses kolom selain NIK/KK, jawab **Ya** pada **Tambah kolom lain?**. Jika sheet tidak punya header NIK/KK, masukkan nomor baris header. Tekan **Spasi** untuk menandai kolom dan **Enter** untuk lanjut. Periksa daftar kolom sebelum mengonfirmasi penyimpanan.

## Mode Langsung

Tanpa dialog, mode `clean` hanya memproses kolom NIK/KK. Contoh dengan workbook dummy:

```console
$ ./dps-excel clean "contoh.xlsx"
3 sel dibersihkan. Hasil: 'contoh_bersih.xlsx'
```

Untuk menentukan nama hasil atau memeriksa versi:

```sh
./dps-excel clean "contoh.xlsx" -o "hasil.xlsx"
./dps-excel --version
```

## Akses dari Mana Saja

Untuk menjalankan `dps-excel` dari folder mana pun, jalankan ini dari root proyek jika `~/.local/bin` sudah ada di `PATH`:

```sh
mkdir -p "$HOME/.local/bin"
ln -s "$(pwd)/dps-excel" "$HOME/.local/bin/dps-excel"
```

Launcher membaca source langsung. Tidak perlu reinstall saat kode berubah. Path relatif mengacu pada folder terminal saat perintah dijalankan.

## Batasan Excel

- Hanya mendukung `.xlsx`. Karakter yang dihapus hanya `U+200C`; angka, rumus, dan kolom yang tidak dipilih dilewati.
- Deteksi otomatis memakai baris pertama yang memuat `NIK`, `No NIK`, `No. NIK`, `No KK`, `No. KK`, `Nomor KK`, atau `Nomor Kartu Keluarga`. Kapitalisasi dan spasi berlebih diabaikan. Pembersihan dimulai dari baris berikutnya.
- NIK/KK yang dibersihkan tetap disimpan sebagai teks, sehingga nol di depan tidak hilang. Digit yang sudah hilang karena Excel menyimpan nomor sebagai angka tidak bisa dipulihkan.
- Hasil default adalah `<nama>_bersih.xlsx`. Jika tidak ada kolom yang cocok atau dipilih, tidak ada file hasil.

## Pengembangan

Lihat [changelog](CHANGELOG.md) dan [tag versi](https://github.com/digitalpencaksilat/dps-excel/tags).

GitHub Actions menjalankan tes setiap kali ada push ke `main`. Jika tes lulus, `semantic-release` menentukan versi berikutnya dari commit sejak tag terakhir: `fix:` untuk patch, `feat:` untuk minor, dan footer `BREAKING CHANGE:` untuk major. Commit `docs:` atau `chore:` biasa tidak memicu rilis sendiri.

Jika ada versi baru, workflow memperbarui `VERSION` (sumber `--version`) dan `CHANGELOG.md`, lalu membuat commit rilis, tag `vX.Y.Z`, dan GitHub Release tanpa PR rilis. Token `GITHUB_TOKEN` harus diizinkan menulis ke `main`; aturan proteksi branch yang menolak commit langsung dapat menghalangi rilis. Setelah rilis, gunakan `git pull --ff-only` untuk mengambil commit versi terbaru.

```sh
.venv/bin/python -m unittest discover -s tests -v
```
