# Monitoring Kendala OSS — DPMPTSP Kota Tegal

Paket frontend statis untuk dihosting melalui **GitHub Pages**. Logo dan favicon sudah dipisahkan dari HTML ke folder `assets/`, sehingga tidak lagi ditanam sebagai data Base64 di dalam `index.html`.

## Struktur folder

```text
monitoring-kendala-oss-github/
├── index.html
├── README.md
└── assets/
    ├── logo-dpmptsp.png
    ├── logo-kota-tegal.png
    ├── logo-oss.png
    └── favicon.png
```

## Cara hosting di GitHub Pages

1. Ekstrak paket ZIP ini.
2. Buat repository baru di GitHub, misalnya `monitoring-kendala-oss`.
3. Unggah **isi folder** paket ini ke direktori utama repository (pastikan `index.html` berada di root, bukan di dalam folder tambahan).
4. Buka **Settings → Pages** pada repository.
5. Pada bagian **Build and deployment**, pilih **Deploy from a branch**.
6. Pilih branch `main` dan folder `/(root)`, lalu klik **Save**.
7. Tunggu proses deployment selesai. URL situs akan ditampilkan pada halaman Pages.

## Catatan

- Nama dan lokasi file aset harus dipertahankan, karena `index.html` merujuk ke file di folder `assets/`.
- Pastikan semua file di folder `assets/` ikut diunggah ke GitHub.
- Paket ini memisahkan frontend dan aset gambar. Jika aplikasi menggunakan Google Apps Script/API sebagai backend, pastikan URL backend dan izin aksesnya tetap sesuai konfigurasi yang digunakan.
- Untuk pembaruan, edit `index.html` atau file di `assets/`, lalu commit/push perubahan ke branch yang dipublikasikan.
