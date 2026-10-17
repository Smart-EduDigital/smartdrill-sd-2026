# 🚀 SmartDrill SD 2026 - Platform Try Out Digital & Bank Soal Interaktif

**SmartDrill SD 2026** adalah produk digital pendidikan berkualitas tinggi yang dirancang untuk persiapan Tes Kemampuan Akademik (TKA) Sekolah Dasar berdasarkan Kurikulum Merdeka dan K-13.

---

## 🌟 Fitur Utama
- **200 Butir Soal Berkualitas:** Terbagi dalam 5 Paket Matematika dan 5 Paket Bahasa Indonesia (masing-masing 20 soal per paket).
- **Format Standar Quizizz (.xlsx):** Siap diimpor langsung ke Quizizz untuk pengelolaan kelas dan ujian berbasis online.
- **Web Try Out Mandiri (`index.html`):** Aplikasi web interaktif lengkap dengan:
  - Validasi Nama & PIN/Token Akses Siswa.
  - Pilihan Mata Pelajaran dan Paket Soal.
  - *Quiz Runner* dengan **Timer Mundur (Countdown Timer)** per soal.
  - Penilaian otomatis (*skor instan*).
  - Menu **Pembahasan & Analisis Jawaban** lengkap untuk evaluasi belajar siswa.
- **100% Gratis & Tanpa Hosting Berbayar:** Dijalankan menggunakan **GitHub Pages**.

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
