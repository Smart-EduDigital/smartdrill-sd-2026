# SmartDrill SD 2026 - Web Try Out Interaktif

Aplikasi web try out digital mandiri untuk produk digital **SmartDrill SD 2026**.

## Struktur File di Folder Ini:
- `index.html` : Tampilan antarmuka web interaktif (Login siswa, pilih paket, runner soal, timer, penilaian otomatis, dan pembahasan lengkap).
- `questions.json` : Bank soal (80 soal lengkap Matematika & Bahasa Indonesia Paket 1 & 2 beserta kunci jawaban dan pembahasan).
- File `.xlsx` : File spreadsheet standar Quizizz asli.

## Cara Menjalankan Secara Lokal (Offline):
1. Pastikan file `index.html` dan `questions.json` berada di dalam satu folder yang sama.
2. Cukup **double-click** (klik dua kali) pada file `index.html` untuk membukanya di browser web Anda (Chrome, Edge, Firefox).

## Cara Upload ke GitHub Pages (Hosting Gratis 100%):
1. Buat akun di [github.com](https://github.com) jika belum punya.
2. Buat repository baru (misal: `smartdrill-sd-2026`), pilih **Public**.
3. Upload kedua file (`index.html` dan `questions.json`) ke repository tersebut.
4. Masuk ke tab **Settings** > **Pages**.
5. Pada bagian **Build and deployment**, pilih Branch `main` (atau `master`) dan folder `/ (root)`, lalu klik **Save**.
6. Dalam 1 menit, web try out Anda sudah online dan bisa diakses via link publik (misal: `https://username.github.io/smartdrill-sd-2026/`).
