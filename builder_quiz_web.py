import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

TARGET_DIR = r"C:\Users\MIP-USER\Documents\Project Soal"

# ==========================================
# 1. DEFINISI DATA SOAL LENGKAP (80 BUTIR)
# ==========================================

# ------------------------------------------
# A. MATEMATIKA PAKET 1
# ------------------------------------------
matematika_p1 = [
    {
        "number": 1,
        "materi": "Operasi Hitung Campuran Bilangan Cacah",
        "question": "Koperasi sekolah memiliki persediaan 35 pak buku tulis. Setiap pak berisi 10 buku tulis. Hari ini koperasi mendapat kiriman tambahan sebanyak 15 pak buku tulis yang masing-masing juga berisi 10 buku. Seluruh buku tulis tersebut akan dikemas ulang ke dalam 25 kardus dengan jumlah yang sama banyak untuk dibagikan ke kelas-kelas. Berapa banyak buku tulis dalam setiap kardus?",
        "options": ["15 buku", "20 buku", "25 buku", "30 buku"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Total persediaan awal = 35 pak × 10 buku = 350 buku.\n2. Tambahan kiriman = 15 pak × 10 buku = 150 buku.\n3. Jumlah seluruh buku = 350 + 150 = 500 buku.\n4. Buku dibagikan ke dalam 25 kardus sama banyak:\n   Banyak buku per kardus = 500 : 25 = 20 buku.\nJadi, jawaban yang benar adalah B (20 buku)."
    },
    {
        "number": 2,
        "materi": "Operasi Bilangan Bulat Negatif (Konteks Suhu)",
        "question": "Suhu daging beku di dalam mesin pendingin (freezer) mula-mula adalah -12°C. Karena terjadi pemadaman listrik selama 40 menit, suhu di dalam mesin pendingin tersebut naik rata-rata 3°C setiap 8 menit. Berapakah suhu daging beku tersebut setelah listrik padam selama 40 menit?",
        "options": ["-3°C", "0°C", "3°C", "6°C"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Lama pemadaman listrik = 40 menit.\n2. Frekuensi kenaikan suhu = 40 menit : 8 menit = 5 kali.\n3. Besar total kenaikan suhu = 5 × 3°C = 15°C.\n4. Suhu akhir daging = -12°C + 15°C = 3°C.\nJadi, suhu daging beku sekarang adalah 3°C (Pilihan C)."
    },
    {
        "number": 3,
        "materi": "Operasi Penjumlahan dan Pengurangan Pecahan Campuran",
        "question": "Ibu Nani memiliki persediaan tepung terigu sebanyak 4 1/2 kg. Ibu menggunakannya untuk membuat kue bolu sebanyak 2 1/4 kg dan membuat donat sebanyak 1 1/8 kg. Berapa kilogram sisa tepung terigu Ibu Nani sekarang?",
        "options": ["1 1/8 kg", "1 1/4 kg", "1 3/8 kg", "1 1/2 kg"],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\nSisa tepung = 4 1/2 - 2 1/4 - 1 1/8\nSamakan penyebutnya menjadi per delapan (8):\n= 4 4/8 - 2 2/8 - 1 1/8\n= (4 - 2 - 1) + (4/8 - 2/8 - 1/8)\n= 1 + 1/8 = 1 1/8 kg.\nAtau dalam pecahan biasa: 36/8 - 18/8 - 9/8 = 9/8 = 1 1/8 kg.\nJadi, pilihan yang tepat adalah A (1 1/8 kg)."
    },
    {
        "number": 4,
        "materi": "Mengurutkan Berbagai Bentuk Pecahan",
        "question": "Perhatikan empat bilangan pecahan berikut: 0,65 ; 3/4 ; 60% ; 7/10. Urutan pecahan dari yang nilainya paling kecil ke paling besar adalah....",
        "options": [
            "60% ; 0,65 ; 7/10 ; 3/4",
            "60% ; 7/10 ; 0,65 ; 3/4",
            "0,65 ; 60% ; 3/4 ; 7/10",
            "7/10 ; 60% ; 0,65 ; 3/4"
        ],
        "correct": "A",
        "explanation": "Langkah penyelesaian (ubah semua ke bentuk desimal):\n- 60% = 60/100 = 0,60\n- 0,65 = 0,65\n- 7/10 = 0,70\n- 3/4 = 0,75\nUrutan dari terkecil ke terbesar:\n0,60 (60%) < 0,65 < 0,70 (7/10) < 0,75 (3/4).\nJadi, urutannya adalah 60% ; 0,65 ; 7/10 ; 3/4 (Pilihan A)."
    },
    {
        "number": 5,
        "materi": "Kelipatan Persekutuan Terkecil (KPK)",
        "question": "Adit berlatih renang setiap 4 hari sekali, Bayu setiap 6 hari sekali, dan Citra setiap 8 hari sekali. Pada hari Senin, 2 Maret 2026, mereka bertiga berlatih renang bersama untuk pertama kalinya. Pada hari dan tanggal berapakah mereka akan berlatih renang bersama untuk kedua kalinya?",
        "options": [
            "Rabu, 25 Maret 2026",
            "Kamis, 26 Maret 2026",
            "Jumat, 27 Maret 2026",
            "Sabtu, 28 Maret 2026"
        ],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Cari KPK dari 4, 6, dan 8:\n   4 = 2²\n   6 = 2 × 3\n   8 = 2³\n   KPK = 2³ × 3 = 8 × 3 = 24 hari.\n2. Mereka akan bersama 24 hari setelah 2 Maret 2026:\n   Tanggal = 2 + 24 = 26 Maret 2026.\n3. Menentukan hari: 24 : 7 = 3 sisa 3 hari.\n   3 hari setelah Senin adalah Kamis (Senin -> Selasa -> Rabu -> Kamis).\nJadi, mereka berlatih bersama lagi pada hari Kamis, 26 Maret 2026 (Pilihan B)."
    },
    {
        "number": 6,
        "materi": "Faktor Persekutuan Terbesar (FPB)",
        "question": "Panitia bakti sosial sekolah menerima sumbangan berupa 72 kg beras, 48 bungkus minyak goreng, dan 36 kotak susu cair. Seluruh barang tersebut akan dikemas ke dalam beberapa paket bantuan dengan jenis dan jumlah barang yang sama banyak pada setiap paketnya tanpa ada yang tersisa. Berapa banyak paket bantuan terbanyak yang dapat dibuat oleh panitia?",
        "options": ["6 paket", "8 paket", "12 paket", "24 paket"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\nJumlah paket terbanyak diperoleh dengan mencari FPB dari 72, 48, dan 36:\n- 72 = 2³ × 3²\n- 48 = 2⁴ × 3\n- 36 = 2² × 3²\nFPB = 2² × 3 = 4 × 3 = 12.\nSetiap paket berisi: Beras (72:12 = 6 kg), Minyak (48:12 = 4 bks), Susu (36:12 = 3 kotak).\nJadi, paket bantuan terbanyak yang dapat dibuat adalah 12 paket (Pilihan C)."
    },
    {
        "number": 7,
        "materi": "Pangkat Dua dan Akar Pangkat Dua (Geometri)",
        "question": "Sebidang tanah aula pertemuan berbentuk persegi memiliki luas 576 m². Di sekeliling tanah aula tersebut akan dipasang kawat pagar pembatas sebanyak satu putaran penuh. Berapakah panjang kawat yang diperlukan untuk mengelilingi tanah tersebut?",
        "options": ["48 meter", "72 meter", "96 meter", "120 meter"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Luas persegi = sisi × sisi = s²\n   s² = 576 m² -> s = √576 = 24 meter.\n2. Keliling persegi = 4 × sisi\n   Keliling = 4 × 24 m = 96 meter.\nJadi, kawat yang diperlukan adalah 96 meter (Pilihan C)."
    },
    {
        "number": 8,
        "materi": "Pangkat Tiga dan Akar Pangkat Tiga (Volume Kubus)",
        "question": "Sebuah bak penampungan air berbentuk kubus mampu menampung air sampai penuh sebanyak 216.000 liter (dm³). Berapakah kedalaman (panjang rusuk bagian dalam) bak penampungan air tersebut dalam satuan meter?",
        "options": ["4 meter", "5 meter", "6 meter", "8 meter"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Volume kubus = s³\n   216.000 dm³ = s³\n   s = ∛216.000 dm³ = 60 dm.\n2. Konversi dari desimeter (dm) ke meter (m):\n   60 dm = 60 : 10 = 6 meter.\nJadi, kedalaman bak penampungan air tersebut adalah 6 meter (Pilihan C)."
    },
    {
        "number": 9,
        "materi": "Perbandingan Senilai dan Selisih",
        "question": "Perbandingan banyak siswa laki-laki dan siswa perempuan di kelas VI SD Harapan Bangsa adalah 4 : 5. Jika selisih banyak siswa laki-laki dan perempuan adalah 4 orang, berapakah jumlah seluruh siswa di kelas VI tersebut?",
        "options": ["32 orang", "36 orang", "40 orang", "45 orang"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Rasio laki-laki : perempuan = 4 : 5.\n2. Selisih perbandingan = 5 - 4 = 1 bagian.\n3. Diketahui selisih nyata = 4 orang, maka nilai 1 bagian = 4 orang.\n4. Jumlah perbandingan seluruh siswa = 4 + 5 = 9 bagian.\n5. Total seluruh siswa = 9 × 4 orang = 36 orang.\n(Laki-laki = 4 × 4 = 16 orang, Perempuan = 5 × 4 = 20 orang. Selisih = 20 - 16 = 4 orang).\nJadi, jumlah seluruh siswa adalah 36 orang (Pilihan B)."
    },
    {
        "number": 10,
        "materi": "Skala dan Jarak Sebenarnya",
        "question": "Jarak antara Kota Melati dan Kota Cempaka pada sebuah peta adalah 6 cm. Skala yang tertera pada peta tersebut adalah 1 : 400.000. Berapakah jarak sebenarnya antara kedua kota tersebut dalam satuan kilometer?",
        "options": ["18 km", "24 km", "28 km", "32 km"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Jarak Sebenarnya = Jarak pada peta : Skala\n   Jarak Sebenarnya = 6 cm × 400.000 = 2.400.000 cm.\n2. Konversi satuan dari cm ke km (dibagi 100.000):\n   2.400.000 cm = 2.400.000 : 100.000 = 24 km.\nJadi, jarak sebenarnya adalah 24 km (Pilihan B)."
    },
    {
        "number": 11,
        "materi": "Keliling Lingkaran dan Penerapannya",
        "question": "Sebuah taman rekreasi keluarga berbentuk lingkaran memiliki panjang jari-jari 14 meter (π = 22/7). Di sekeliling tepi taman tersebut akan ditanami pohon palem perindang dengan jarak antar-pohon sebesar 4 meter. Berapa banyak pohon palem yang dibutuhkan?",
        "options": ["18 pohon", "20 pohon", "22 pohon", "24 pohon"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Keliling lingkaran = 2 × π × r\n   Keliling = 2 × (22/7) × 14 m\n   Keliling = 2 × 22 × 2 m = 88 meter.\n2. Banyak pohon palem = Keliling taman : jarak antar-pohon\n   Banyak pohon = 88 m : 4 m = 22 pohon.\nJadi, banyak pohon palem yang dibutuhkan adalah 22 pohon (Pilihan C)."
    },
    {
        "number": 12,
        "materi": "Luas Gabungan Bangun Datar",
        "question": "Sebuah bidang hiasan dinding berbentuk gabungan persegi panjang dan segitiga sama kaki. Persegi panjang memiliki panjang 18 cm dan lebar 10 cm. Di atas salah satu sisi panjangnya menempel alas segitiga sama kaki yang berukuran alas 18 cm dan tinggi 8 cm. Berapakah luas total bangun gabungan tersebut?",
        "options": ["216 cm²", "252 cm²", "288 cm²", "324 cm²"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Luas Persegi Panjang = panjang × lebar = 18 cm × 10 cm = 180 cm².\n2. Luas Segitiga = 1/2 × alas × tinggi = 1/2 × 18 cm × 8 cm = 9 × 8 = 72 cm².\n3. Luas Gabungan = Luas persegi panjang + Luas segitiga\n   Luas Total = 180 cm² + 72 cm² = 252 cm².\nJadi, luas gabungan tersebut adalah 252 cm² (Pilihan B)."
    },
    {
        "number": 13,
        "materi": "Volume Bangun Ruang Balok (Pecahan Bagian)",
        "question": "Sebuah bak penampungan air berbentuk balok memiliki ukuran panjang 120 cm, lebar 80 cm, dan tinggi 100 cm. Bak tersebut telah terisi air sebanyak 3/4 bagian. Berapa liter air yang masih harus ditambahkan agar bak penampungan air tersebut terisi penuh?",
        "options": ["180 liter", "240 liter", "320 liter", "720 liter"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Volume total balok = p × l × t\n   V = 120 cm × 80 cm × 100 cm = 960.000 cm³.\n2. Konversi ke liter (dm³):\n   960.000 cm³ = 960.000 : 1.000 = 960 liter.\n3. Karena telah terisi 3/4 bagian, maka bagian yang kosong = 1 - 3/4 = 1/4 bagian.\n4. Air yang harus ditambahkan = 1/4 × 960 liter = 240 liter.\nJadi, air yang harus ditambahkan adalah 240 liter (Pilihan B)."
    },
    {
        "number": 14,
        "materi": "Sifat-sifat Bangun Ruang",
        "question": "Perhatikan ciri-ciri bangun ruang berikut:\n(1) Memiliki 5 buah sisi (1 sisi alas dan 4 sisi tegak berbentuk segitiga).\n(2) Memiliki 8 buah rusuk.\n(3) Memiliki 5 buah titik sudut (satu di antaranya merupakan titik puncak).\nBangun ruang yang memiliki karakteristik tersebut adalah....",
        "options": ["Prisma segitiga", "Limas segitiga", "Limas segi empat", "Prisma segi empat"],
        "correct": "C",
        "explanation": "Analisis sifat bangun ruang:\n- Prisma segitiga memiliki 5 sisi, 9 rusuk, 6 titik sudut.\n- Limas segitiga memiliki 4 sisi, 6 rusuk, 4 titik sudut.\n- Limas segi empat memiliki 1 alas segi empat + 4 sisi tegak segitiga = 5 sisi, 8 rusuk, dan 5 titik sudut (termasuk 1 titik puncak).\nJadi, bangun ruang yang dimaksud adalah Limas segi empat (Pilihan C)."
    },
    {
        "number": 15,
        "materi": "Pengukuran Waktu, Kecepatan, dan Jarak",
        "question": "Pak Rahmat berangkat dari Bandung menuju Jakarta mengendarai mobil dengan kecepatan rata-rata 75 km/jam. Jarak yang ditempuh adalah 150 km. Jika Pak Rahmat berangkat dari Bandung pada pukul 06.30 WIB dan sempat beristirahat di rest area selama 30 menit di tengah perjalanan, pada pukul berapakah Pak Rahmat tiba di Jakarta?",
        "options": ["08.30 WIB", "09.00 WIB", "09.15 WIB", "09.30 WIB"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Waktu tempuh murni = Jarak : Kecepatan\n   Waktu = 150 km : 75 km/jam = 2 jam.\n2. Waktu istirahat = 30 menit (0,5 jam).\n3. Total durasi perjalanan = 2 jam + 30 menit = 2 jam 30 menit.\n4. Waktu kedatangan = Waktu berangkat + Total durasi\n   = 06.30 WIB + 02.30 = 09.00 WIB.\nJadi, Pak Rahmat tiba di Jakarta pada pukul 09.00 WIB (Pilihan B)."
    },
    {
        "number": 16,
        "materi": "Debit Aliran Air",
        "question": "Sebuah kolam renang anak dengan kapasitas volume air 3.600 liter dikosongkan untuk dibersihkan menggunakan pipa pembuangan. Seluruh air di dalam kolam habis terkuras dalam waktu 30 menit. Berapakah debit aliran air yang keluar melalui pipa pembuangan tersebut dalam satuan liter/detik?",
        "options": ["1,5 liter/detik", "2 liter/detik", "2,5 liter/detik", "3 liter/detik"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Volume air = 3.600 liter.\n2. Waktu pengurasan = 30 menit = 30 × 60 detik = 1.800 detik.\n3. Debit = Volume : Waktu\n   Debit = 3.600 liter : 1.800 detik = 2 liter/detik.\nJadi, debit aliran air adalah 2 liter/detik (Pilihan B)."
    },
    {
        "number": 17,
        "materi": "Membaca dan Menafsirkan Data Tabel",
        "question": "Perhatikan tabel hasil panen jagung Desa Makmur berikut:\n- Tahun 2021: 45 ton\n- Tahun 2022: 50 ton\n- Tahun 2023: 65 ton\n- Tahun 2024: 55 ton\n- Tahun 2025: 75 ton\nBerapakah selisih hasil panen jagung tertinggi dan hasil panen jagung terendah selama periode tersebut?",
        "options": ["20 ton", "25 ton", "30 ton", "35 ton"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Hasil panen tertinggi terjadi pada tahun 2025 = 75 ton.\n2. Hasil panen terendah terjadi pada tahun 2021 = 45 ton.\n3. Selisih = 75 ton - 45 ton = 30 ton.\nJadi, selisih hasil panen tertinggi dan terendah adalah 30 ton (Pilihan C)."
    },
    {
        "number": 18,
        "materi": "Membaca Data Diagram Batang (Rata-rata Penjualan)",
        "question": "Data penjualan buku ensiklopedia anak di sebuah toko buku selama enam hari berturut-turut tercatat sebagai berikut: Senin 35 buku, Selasa 45 buku, Rabu 40 buku, Kamis 55 buku, Jumat 50 buku, dan Sabtu 75 buku. Berapakah rata-rata penjualan buku per hari selama enam hari tersebut?",
        "options": ["45 buku", "48 buku", "50 buku", "52 buku"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Jumlah seluruh buku terjual = 35 + 45 + 40 + 55 + 50 + 75 = 300 buku.\n2. Banyaknya hari penjualan = 6 hari.\n3. Rata-rata per hari = 300 : 6 = 50 buku.\nJadi, rata-rata penjualan buku per hari adalah 50 buku (Pilihan C)."
    },
    {
        "number": 19,
        "materi": "Menghitung Nilai Pengganti pada Rata-rata (Mean)",
        "question": "Nilai ulangan matematika dari 8 orang siswa adalah: 75, 80, 85, 70, 90, 85, 75, dan 80. Ketika nilai ulangan milik siswa ke-9 dimasukkan ke dalam kelompok tersebut, nilai rata-rata keseluruhan menjadi 81. Berapakah nilai ulangan matematika siswa ke-9 tersebut?",
        "options": ["86", "87", "88", "89"],
        "correct": "D",
        "explanation": "Langkah penyelesaian:\n1. Jumlah nilai 8 siswa = 75 + 80 + 85 + 70 + 90 + 85 + 75 + 80 = 640.\n2. Nilai rata-rata 9 siswa = 81.\n   Total nilai 9 siswa = 9 × 81 = 729.\n3. Nilai siswa ke-9 = Total 9 siswa - Total 8 siswa\n   Nilai siswa ke-9 = 729 - 640 = 89.\nJadi, nilai ulangan siswa ke-9 adalah 89 (Pilihan D)."
    },
    {
        "number": 20,
        "materi": "Statistika Dasar (Modus dan Median)",
        "question": "Data berat badan (dalam kg) dari 15 siswa kelas VI SD disajikan sebagai berikut:\n32, 34, 32, 36, 35, 34, 32, 38, 35, 36, 32, 32, 35, 37, 34.\nModus dan median dari data berat badan siswa tersebut berturut-turut adalah....",
        "options": [
            "Modus = 32 kg dan Median = 34 kg",
            "Modus = 34 kg dan Median = 35 kg",
            "Modus = 32 kg dan Median = 35 kg",
            "Modus = 35 kg dan Median = 34 kg"
        ],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Menentukan Modus (data paling sering muncul):\n   - 32 muncul sebanyak 5 kali\n   - 34 muncul sebanyak 3 kali\n   - 35 muncul sebanyak 3 kali\n   - 36 muncul sebanyak 2 kali\n   - 37 muncul sebanyak 1 kali\n   - 38 muncul sebanyak 1 kali\n   Modus = 32 kg.\n2. Menentukan Median (urutkan data dari terkecil):\n   32, 32, 32, 32, 32, 34, 34, [34], 35, 35, 35, 36, 36, 37, 38.\n   Data ke-8 (nilai tengah) adalah 34 kg.\n   Median = 34 kg.\nJadi, modus = 32 kg dan median = 34 kg (Pilihan A)."
    }
]

# ------------------------------------------
# B. MATEMATIKA PAKET 2
# ------------------------------------------
matematika_p2 = [
    {
        "number": 1,
        "materi": "Operasi Hitung Campuran Bilangan Cacah",
        "question": "Pak Arman memanen buah apel sebanyak 18 keranjang. Setiap keranjang berisi 45 kg apel. Dari seluruh hasil panen tersebut, sebanyak 110 kg dibagikan kepada tetangga sekitar dan kerabat, sedangkan sisanya dijual kepada 10 orang pedagang buah di pasar dengan jumlah yang sama banyak. Berapa kilogram apel yang diterima oleh masing-masing pedagang?",
        "options": ["65 kg", "70 kg", "75 kg", "80 kg"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Total panen apel = 18 keranjang × 45 kg = 810 kg.\n2. Sisa setelah dibagikan = 810 kg - 110 kg = 700 kg.\n3. Dibagikan ke 10 pedagang sama banyak = 700 : 10 = 70 kg.\nJadi, setiap pedagang menerima 70 kg apel (Pilihan B)."
    },
    {
        "number": 2,
        "materi": "Operasi Bilangan Bulat (Perubahan Suhu Udara)",
        "question": "Suhu udara di kawasan dataran tinggi pegunungan pada dini hari tercatat -4°C. Pada siang hari yang cerah, suhunya mengalami kenaikan sebesar 19°C. Menjelang malam hari, suhu udara kembali turun sebesar 6°C. Berapakah suhu udara di puncak pegunungan tersebut pada malam hari?",
        "options": ["7°C", "9°C", "11°C", "13°C"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Suhu awal = -4°C.\n2. Siang hari naik 19°C -> -4°C + 19°C = 15°C.\n3. Malam hari turun 6°C -> 15°C - 6°C = 9°C.\nPersamaan utuh: -4 + 19 - 6 = 9°C.\nJadi, suhu udara pada malam hari adalah 9°C (Pilihan B)."
    },
    {
        "number": 3,
        "materi": "Operasi Hitung Pecahan Campuran",
        "question": "Sebuah jeriken berisi minyak goreng sebanyak 5 1/4 liter. Ibu menggunakannya untuk menggoreng kerupuk sebanyak 2 1/2 liter. Kemudian, Ayah membeli minyak kemasan sebanyak 1 3/4 liter dan menuangkannya ke dalam jeriken tersebut. Berapa liter volume minyak goreng di dalam jeriken sekarang?",
        "options": ["4 1/4 liter", "4 1/2 liter", "4 3/4 liter", "5 liter"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\nVolume minyak = 5 1/4 - 2 1/2 + 1 3/4\nSamakan penyebut ke per empat (4):\n= 5 1/4 - 2 2/4 + 1 3/4\nUbah ke pecahan biasa:\n= 21/4 - 10/4 + 7/4\n= (21 - 10 + 7) / 4 = 18/4 = 4 2/4 = 4 1/2 liter (atau 4,5 liter).\nJadi, volume minyak sekarang adalah 4 1/2 liter (Pilihan B)."
    },
    {
        "number": 4,
        "materi": "Membandingkan dan Mengurutkan Pecahan",
        "question": "Di laboratorium sains sekolah terdapat empat botol larutan dengan tingkat kepekatan zat aktif sebagai berikut: Botol P = 45%, Botol Q = 0,55, Botol R = 3/5, dan Botol S = 0,4. Urutan botol larutan dari yang memiliki kadar kepekatan paling tinggi ke paling rendah adalah....",
        "options": [
            "Botol R, Botol Q, Botol P, Botol S",
            "Botol R, Botol P, Botol Q, Botol S",
            "Botol Q, Botol R, Botol P, Botol S",
            "Botol S, Botol P, Botol Q, Botol R"
        ],
        "correct": "A",
        "explanation": "Langkah penyelesaian (konversi ke desimal):\n- Botol P: 45% = 0,45\n- Botol Q: 0,55\n- Botol R: 3/5 = 6/10 = 0,60\n- Botol S: 0,4 = 0,40\nUrutan dari tertinggi ke terendah:\n0,60 (R) > 0,55 (Q) > 0,45 (P) > 0,40 (S).\nJadi, urutannya adalah Botol R, Botol Q, Botol P, Botol S (Pilihan A)."
    },
    {
        "number": 5,
        "materi": "Kelipatan Persekutuan Terkecil (KPK Tiga Bilangan)",
        "question": "Tiga buah bel alarm penanda pergantian kegiatan di asrama sekolah berbunyi secara berkala. Bel A berbunyi setiap 15 menit, bel B setiap 20 menit, dan bel C setiap 30 menit. Jika ketiga bel berbunyi bersamaan pada pukul 07.15 WIB, pada pukul berapakah ketiga bel tersebut akan berbunyi bersamaan lagi untuk kedua kalinya?",
        "options": ["07.45 WIB", "08.00 WIB", "08.15 WIB", "08.30 WIB"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Cari KPK dari 15, 20, dan 30:\n   15 = 3 × 5\n   20 = 2² × 5\n   30 = 2 × 3 × 5\n   KPK = 2² × 3 × 5 = 4 × 15 = 60 menit (1 jam).\n2. Waktu berbunyi bersama berikutnya = 07.15 WIB + 1 jam = 08.15 WIB.\nJadi, ketiga bel berbunyi bersamaan lagi pada pukul 08.15 WIB (Pilihan C)."
    },
    {
        "number": 6,
        "materi": "Faktor Persekutuan Terbesar (FPB Tiga Bilangan)",
        "question": "Kelompok Tani Makmur memiliki persediaan 60 bibit pohon cabai, 75 bibit tomat, dan 90 bibit terung. Seluruh bibit tanaman tersebut akan dibagikan kepada beberapa kelompok warga pemukiman dengan komposisi jumlah dan jenis bibit yang sama banyak pada setiap kelompoknya tanpa ada bibit yang tersisa. Berapa banyak kelompok warga paling banyak yang dapat menerima bibit tersebut?",
        "options": ["10 kelompok", "12 kelompok", "15 kelompok", "20 kelompok"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\nCari FPB dari 60, 75, dan 90:\n- 60 = 2² × 3 × 5\n- 75 = 3 × 5²\n- 90 = 2 × 3² × 5\nFPB = 3 × 5 = 15.\nMasing-masing kelompok menerima:\nCabai = 60:15 = 4 bibit, Tomat = 75:15 = 5 bibit, Terung = 90:15 = 6 bibit.\nJadi, paling banyak ada 15 kelompok warga yang dapat menerima (Pilihan C)."
    },
    {
        "number": 7,
        "materi": "Akar Pangkat Dua dalam Keliling Kebun Persegi",
        "question": "Pak Santo memiliki kebun sayur berbentuk persegi dengan luas 784 m². Di sekeliling kebun tersebut akan dipasang tiang pancang bambu penahan pagar dengan jarak antar-tiang sejauh 2 meter. Berapakah banyak tiang pancang bambu yang dibutuhkan Pak Santo?",
        "options": ["48 buah", "52 buah", "56 buah", "64 buah"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Sisi persegi = √Luas = √784 m² = 28 meter.\n2. Keliling kebun persegi = 4 × sisi = 4 × 28 m = 112 meter.\n3. Banyak tiang bambu = Keliling : jarak antar-tiang\n   Banyak tiang = 112 m : 2 m = 56 buah.\nJadi, tiang bambu yang dibutuhkan adalah 56 buah (Pilihan C)."
    },
    {
        "number": 8,
        "materi": "Akar Pangkat Tiga dan Volume Kubus",
        "question": "Sebuah akuarium kaca berbentuk kubus terisi air tepat separuh (1/2 bagian). Volume air di dalam akuarium tersebut tercatat 62,5 liter (dm³). Berapakah panjang rusuk akuarium tersebut dalam satuan sentimeter (cm)?",
        "options": ["40 cm", "50 cm", "60 cm", "70 cm"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Karena 1/2 volume = 62,5 liter, maka:\n   Volume penuh kubus = 62,5 × 2 = 125 liter (dm³).\n2. Panjang rusuk kubus = ∛Volume = ∛125 dm³ = 5 dm.\n3. Konversi dari dm ke cm (dikalikan 10):\n   5 dm = 5 × 10 = 50 cm.\nJadi, panjang rusuk akuarium tersebut adalah 50 cm (Pilihan B)."
    },
    {
        "number": 9,
        "materi": "Rasio Umur dan Penjumlahan",
        "question": "Perbandingan umur Dimas dan Kakaknya adalah 3 : 7. Jika umur Kakak saat ini adalah 28 tahun, berapakah jumlah umur Dimas dan Kakaknya?",
        "options": ["36 tahun", "40 tahun", "42 tahun", "44 tahun"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Rasio Dimas : Kakak = 3 : 7.\n2. Umur Kakak = 28 tahun (7 bagian).\n   Nilai 1 bagian = 28 : 7 = 4 tahun.\n3. Umur Dimas = 3 bagian × 4 tahun = 12 tahun.\n4. Jumlah umur = 12 tahun + 28 tahun = 40 tahun.\n(Atau: (3 + 7) × 4 = 10 × 4 = 40 tahun).\nJadi, jumlah umur mereka adalah 40 tahun (Pilihan B)."
    },
    {
        "number": 10,
        "materi": "Skala Denah dan Luas Sebenarnya",
        "question": "Denah gedung serbaguna digambar dengan skala 1 : 200. Pada gambar denah tersebut, ukuran panjang gedung adalah 15 cm dan lebarnya 10 cm. Berapakah luas sebenarnya dari lantai gedung serbaguna tersebut dalam satuan meter persegi (m²)?",
        "options": ["300 m²", "450 m²", "600 m²", "750 m²"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Panjang sebenarnya = 15 cm × 200 = 3.000 cm = 30 meter.\n2. Lebar sebenarnya = 10 cm × 200 = 2.000 cm = 20 meter.\n3. Luas sebenarnya = Panjang sebenarnya × Lebar sebenarnya\n   Luas = 30 m × 20 m = 600 m².\nJadi, luas lantai sebenarnya adalah 600 m² (Pilihan C)."
    },
    {
        "number": 11,
        "materi": "Luas Lingkaran",
        "question": "Sebuah taman kota berbentuk lingkaran memiliki diameter 28 meter (π = 22/7). Seluruh permukaan tanah taman tersebut akan ditanami rumput hias gajah mini. Berapakah luas tanah taman yang akan ditanami rumput tersebut?",
        "options": ["588 m²", "616 m²", "644 m²", "672 m²"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Diameter (d) = 28 meter, maka jari-jari (r) = 28 : 2 = 14 meter.\n2. Luas lingkaran = π × r²\n   Luas = (22/7) × 14 m × 14 m\n   Luas = 22 × 2 × 14 = 44 × 14 = 616 m².\nJadi, luas tanah yang ditanami rumput adalah 616 m² (Pilihan B)."
    },
    {
        "number": 12,
        "materi": "Luas Trapesium Siku-siku",
        "question": "Sebidang tanah pekarangan rumah berbentuk trapesium siku-siku memiliki panjang dua sisi sejajar masing-masing 16 meter dan 24 meter. Jarak tegak lurus antara kedua sisi sejajar tersebut (tinggi trapesium) adalah 12 meter. Berapakah luas pekarangan tersebut?",
        "options": ["200 m²", "240 m²", "280 m²", "320 m²"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\nLuas trapesium = 1/2 × (jumlah sisi sejajar) × tinggi\nLuas = 1/2 × (16 m + 24 m) × 12 m\nLuas = 1/2 × 40 m × 12 m = 20 × 12 = 240 m².\nJadi, luas pekarangan rumah tersebut adalah 240 m² (Pilihan B)."
    },
    {
        "number": 13,
        "materi": "Volume Tabung (Silinder)",
        "question": "Sebuah tandon penampung air hujan berbentuk tabung memiliki jari-jari alas 70 cm dan tinggi 1,5 meter (150 cm). Tandon tersebut diisi air sampai penuh (gunakan π = 22/7). Berapa liter volume air di dalam tandon tersebut?",
        "options": ["1.840 liter", "2.100 liter", "2.310 liter", "2.520 liter"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Jari-jari (r) = 70 cm, tinggi (t) = 150 cm.\n2. Volume tabung = π × r² × t\n   V = (22/7) × 70 × 70 × 150 cm³\n   V = 22 × 10 × 70 × 150 = 220 × 10.500 = 2.310.000 cm³.\n3. Konversi ke liter (dm³):\n   2.310.000 cm³ : 1.000 = 2.310 liter.\nJadi, volume air adalah 2.310 liter (Pilihan C)."
    },
    {
        "number": 14,
        "materi": "Karakteristik Bangun Ruang Sisi Lengkung (Tabung)",
        "question": "Perhatikan sifat-sifat bangun ruang berikut:\n(1) Memiliki 2 sisi datar berbentuk lingkaran yang kongruen (alas dan tutup).\n(2) Memiliki 1 sisi lengkung yang menghubungkan alas dan tutup (selimut).\n(3) Memiliki 2 buah rusuk lengkung dan tidak memiliki titik sudut.\nBangun ruang yang memiliki sifat-sifat tersebut adalah....",
        "options": ["Kerucut", "Bola", "Tabung", "Prisma lingkaran"],
        "correct": "C",
        "explanation": "Analisis sifat:\n- Kerucut: memiliki 1 titik puncak, 1 sisi lengkung, 1 alas lingkaran.\n- Bola: hanya memiliki 1 sisi lengkung tanpa rusuk.\n- Tabung: memiliki 3 sisi (2 lingkaran sejajar kongruen + 1 selimut lengkung), 2 rusuk lengkung, dan 0 titik sudut.\nJadi, bangun ruang yang dimaksud adalah Tabung (Pilihan C)."
    },
    {
        "number": 15,
        "materi": "Kecepatan Rata-rata dan Waktu Kedatangan",
        "question": "Kereta Api Logawa melaju dari Stasiun Purwokerto menuju Stasiun Surabaya dengan kecepatan rata-rata 80 km/jam. Jarak rute rel yang ditempuh adalah 240 km. Jika kereta api berangkat pada pukul 08.20 WIB dan sempat berhenti transit di stasiun antara selama 25 menit, pada pukul berapakah kereta tiba di Stasiun Surabaya?",
        "options": ["11.20 WIB", "11.35 WIB", "11.45 WIB", "12.00 WIB"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Waktu berjalan murni = Jarak : Kecepatan\n   Waktu = 240 km : 80 km/jam = 3 jam.\n2. Waktu transit = 25 menit.\n3. Total waktu perjalanan = 3 jam + 25 menit = 3 jam 25 menit.\n4. Jam tiba = 08.20 WIB + 03.25 = 11.45 WIB.\nJadi, kereta tiba di Surabaya pada pukul 11.45 WIB (Pilihan C)."
    },
    {
        "number": 16,
        "materi": "Debit dan Waktu Pengosongan",
        "question": "Sebuah tandon air berkapasitas 1.500 liter terisi penuh. Di bagian bawah tandon dipasang kran pembuangan air dengan debit aliran konstan 25 liter/menit. Berapa menit waktu yang dibutuhkan untuk mengosongkan seluruh air di dalam tandon tersebut?",
        "options": ["50 menit", "55 menit", "60 menit", "75 menit"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\nWaktu = Volume : Debit\nWaktu = 1.500 liter : 25 liter/menit = 60 menit (setara 1 jam).\nJadi, waktu yang dibutuhkan adalah 60 menit (Pilihan C)."
    },
    {
        "number": 17,
        "materi": "Membaca Data Penjualan Toko",
        "question": "Data kunjungan pemustaka di Perpustakaan Daerah selama 5 hari tercatat sebagai berikut: Senin 60 orang, Selasa 75 orang, Rabu 90 orang, Kamis 65 orang, dan Jumat 110 orang. Berapakah jumlah total kunjungan pemustaka selama 5 hari tersebut?",
        "options": ["380 orang", "390 orang", "400 orang", "410 orang"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\nTotal kunjungan = 60 + 75 + 90 + 65 + 110 = 400 orang.\nJadi, jumlah total kunjungan adalah 400 orang (Pilihan C)."
    },
    {
        "number": 18,
        "materi": "Menafsirkan Diagram Lingkaran Persentase",
        "question": "Sebuah survei kegiatan ekstrakurikuler diikuti oleh 120 siswa kelas VI dengan rincian persentase: Sepak Bola 35%, Bulu Tangkis 25%, Renang 20%, dan Seni Musik sisanya. Berapakah selisih banyak siswa yang memilih ekstrakurikuler Sepak Bola dengan Bulu Tangkis?",
        "options": ["10 orang", "12 orang", "15 orang", "18 orang"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Siswa Sepak Bola = 35% × 120 = (35/100) × 120 = 42 orang.\n2. Siswa Bulu Tangkis = 25% × 120 = (25/100) × 120 = 30 orang.\n3. Selisih = 42 orang - 30 orang = 12 orang.\n(Atau: (35% - 25%) × 120 = 10% × 120 = 12 orang).\nJadi, selisihnya adalah 12 orang (Pilihan B)."
    },
    {
        "number": 19,
        "materi": "Menghitung Nilai Pengganti Rata-rata",
        "question": "Nilai rata-rata ulangan harian IPA dari 9 orang siswa adalah 78. Ketika nilai ulangan milik Bayu digabungkan ke dalam kelompok tersebut, nilai rata-ratanya naik menjadi 80. Berapakah nilai ulangan IPA Bayu?",
        "options": ["92", "95", "98", "100"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Jumlah nilai 9 siswa = 9 × 78 = 702.\n2. Jumlah nilai setelah Bayu bergabung (10 siswa) = 10 × 80 = 800.\n3. Nilai Bayu = Total 10 siswa - Total 9 siswa\n   Nilai Bayu = 800 - 702 = 98.\nJadi, nilai ulangan Bayu adalah 98 (Pilihan C)."
    },
    {
        "number": 20,
        "materi": "Statistika Dasar (Modus dan Median Genap)",
        "question": "Nilai ulangan matematika dari 12 orang siswa adalah sebagai berikut:\n70, 80, 75, 90, 85, 80, 75, 80, 85, 95, 80, 70.\nModus dan median dari data nilai ulangan tersebut adalah....",
        "options": [
            "Modus = 80 dan Median = 80",
            "Modus = 75 dan Median = 80",
            "Modus = 80 dan Median = 77,5",
            "Modus = 85 dan Median = 80"
        ],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Menentukan Modus (nilai paling sering muncul):\n   - Nilai 70 muncul 2 kali\n   - Nilai 75 muncul 2 kali\n   - Nilai 80 muncul 4 kali\n   - Nilai 85 muncul 2 kali\n   - Nilai 90 muncul 1 kali\n   - Nilai 95 muncul 1 kali\n   Modus = 80.\n2. Menentukan Median (urutkan 12 data dari terkecil):\n   70, 70, 75, 75, 80, [80, 80], 80, 85, 85, 90, 95.\n   Nilai tengah berada di antara data ke-6 (80) dan data ke-7 (80).\n   Median = (80 + 80) : 2 = 80.\nJadi, modus = 80 dan median = 80 (Pilihan A)."
    }
]

# ------------------------------------------
# C. BAHASA INDONESIA PAKET 1
# ------------------------------------------
teks_bi_p1_teks1 = """Bacalah teks informasi berikut dengan cermat untuk menjawab soal nomor 1 sampai 5!

Penyu Sisik: Sang Penjaga Terumbu Karang Nusantara

Penyu sisik (Eretmochelys imbricata) merupakan salah satu spesies penyu laut yang memiliki peran sangat krusial dalam menjaga keseimbangan ekosistem terumbu karang di perairan tropis Indonesia. Ciri khas utama satwa ini terletak pada bentuk paruhnya yang runcing menyerupai paruh burung elang serta susunan sisik karapasnya yang bertumpuk indah keemasan. Makanan utama penyu sisik adalah spons laut (hewan porifera) yang hidup menempel pada karang. Dengan memakan spons laut tersebut, penyu sisik mengontrol populasinya agar tidak tumbuh berlebihan dan menutupi karang. Tindakan alami ini memberi ruang bagi terumbu karang untuk tumbuh subur dan menjadi rumah serta tempat berkembang biak bagi aneka jenis ikan.

Namun, keberadaan penyu sisik saat ini berada dalam bayang-bayang ancaman kepunahan yang nyata. Populasi satwa langka ini terus mengalami penurunan tajam akibat perburuan liar untuk diambil cangkangnya sebagai bahan kerajinan bernilai tinggi, perdagangan telur penyu secara ilegal, serta pencemaran sampah plastik di lautan. Penyu kerap mengira kantong plastik transparan yang terapung di perairan sebagai ubur-ubur santapan mereka. Akibatnya, plastik menyumbat saluran pencernaan dan menyebabkan kematian tragis bagi ribuan penyu setiap tahunnya.

Untuk mencegah kepunahan satwa karismatik ini, pemerintah bersama komunitas pelestari alam menggiatkan berbagai upaya konservasi dan penetasan semi-alami, seperti di pesisir Sukamade, Jawa Timur dan Kepulauan Derawan, Kalimantan Timur. Di pusat konservasi, sarang telur penyu di pantai dijaga ketat dari incaran predator maupun pemburu. Telur-telur yang berhasil diamankan dipindahkan ke bak penetasan khusus hingga menetas. Setelah menetas, anak penyu (tukik) dirawat secara intensif selama beberapa hari sebelum akhirnya dilepasliarkan kembali ke laut lepas pada malam atau dini hari guna menghindari serangan burung pemangsa."""

teks_bi_p1_teks2 = """Bacalah cerita anak berikut dengan cermat untuk menjawab soal nomor 6 sampai 10!

Kejujuran di Balik Kantong Belacu Kakek Suro

Siang itu, terik matahari membakar jalanan Desa Sukamaju. Arga mengayuh sepeda ontel tuanya dengan peluh yang mengucur di pelipis. Di keranjang depan sepeda, tersusun rapi beberapa ikat bayam dan kangkung segar hasil kebun pekarangan rumahnya yang hendak ia jajakan ke pasar desa demi membantu keperluan dapur ibunya. Ketika melewati tikungan jembatan kayu di perbatasan desa, roda depan sepedanya menabrak sesuatu. Arga mengerem mendadak dan menoleh ke bawah. Di tepi semak-semak rumput gajah, tergeletak sebuah kantong kain belacu berwarna cokelat kusam.

Arga turun dari sepeda dan memungut kantong tersebut. Alangkah terkejutnya ia ketika membuka resletingnya: di dalamnya terdapat beberapa lembar uang ratusan ribu rupiah yang diikat karet gelang, sebuah kartu identitas penduduk, dan buku pemeriksaan kesehatan lansia atas nama Kakek Suro. Kakek Suro adalah seorang perajin anyaman bambu tua yang tinggal sebatang kara di ujung dusun.

Di dalam hati Arga, sempat terlintas bisikan godaan. Sepatu sekolahnya sudah jebol di bagian tumit, dan uang itu lebih dari cukup untuk membeli sepatu baru yang bagus. Namun, seketika petuah bijak sang ibu bergaung kuat di dadanya: "Kejujuran adalah kehormatan hidup kita, Arga. Mengambil hak orang lain hanya akan menanam duri dalam jiwa." 

Tanpa ragu lagi, Arga memutar haluan sepedanya menuju gubuk Kakek Suro. Sesampainya di sana, Arga mendapati sang kakek sedang duduk lemas di tangga bambu sambil meneteskan air mata. Ternyata, uang tersebut adalah tabungan hasil penjualan tampah dan bakul anyaman selama empat bulan yang rencananya akan dipakai untuk menebus obat asma di puskesmas. Ketika Arga mengulurkan kantong belacu itu utuh tanpa berkurang sedikit pun, Kakek Suro seketika memeluk Arga erat dengan linangan air mata haru dan tak henti memanjatkan doa berkah bagi anak yang berbudi luhur itu."""

teks_bi_p1_teks3 = """Bacalah kutipan dialog wawancara berikut untuk menjawab soal nomor 11 sampai 14!

Wawancara Cilik: Menggugah Kepedulian Lewat Bank Sampah Sekolah

Dika (Pewawancara) : "Selamat pagi, Ibu Ratna. Kami dari tim buletin sekolah ingin menggali informasi tentang keberhasilan program 'Bank Sampah Berseri' yang Ibu bimbing. Sebenarnya apa tujuan utama dibentuknya bank sampah ini di sekolah kita?"
Ibu Ratna (Pembina) : "Selamat pagi, Dika. Tujuan utamanya bukan semata-mata mengumpulkan sampah, melainkan membangun kesadaran dan kebiasaan berwawasan lingkungan hidup sejak dini. Kami ingin para siswa memahami konsep pilah sampah: mana sampah organik yang bisa diolah jadi kompos, mana anorganik yang bisa didaur ulang bernilai ekonomis, dan mana residu."
Dika : "Bagaimana cara kerja penyetoran sampah oleh siswa setiap pekannya, Bu?"
Ibu Ratna : "Setiap hari Jumat pagi, siswa membawa sampah anorganik kering dari rumah maupun sekolah, seperti botol plastik, kardus bekas, atau kaleng minuman. Pengurus cilik bank sampah akan menimbang beratnya, mencatatnya, lalu mengonversikannya menjadi saldo tabungan di buku rekening siswa. Saldo tabungan ini pada akhir semester dapat ditukarkan dengan buku tulis, pensil warna, atau bibit tanaman hias."
Dika : "Sungguh inovatif! Lalu apa dampak paling signifikan yang dirasakan sekolah setelah program ini berjalan satu semester, Bu?"
Ibu Ratna : "Alhamdulillah, volume timbulan sampah sekolah yang dibuang ke TPA berkurang drastis hingga 60%. Selain itu, ruang kelas menjadi sangat bersih dan siswa terbiasa bertanggung jawab terhadap lingkungan sekitarnya." """

teks_bi_p1_teks4 = """Bacalah petunjuk prosedur berikut untuk menjawab soal nomor 15 dan 16!

Petunjuk Praktis Menanam Sayuran Hidroponik Sistem Sumbu (Wick System)
(1) Siapkan botol plastik bekas air mineral ukuran 1,5 liter, lalu potong secara horizontal menjadi dua bagian menggunakan gunting atau *cutter*.
(2) Buatlah lubang kecil pada tutup botol menggunakan solder atau paku panas, lalu selipkan sehelai kain flanel sepanjang 15 cm sebagai sumbu kapiler.
(3) Pasang bagian atas botol yang sudah bersumbu ke dalam potongan bagian bawah botol secara terbalik (leher botol menghadap ke dasar).
(4) [...]
(5) Masukkan bibit tanaman sawi atau selada yang telah disemai (berdaun 3 sampai 4 helai) dengan hati-hati ke dalam media tanam tersebut.
(6) Tuangkan larutan nutrisi hidroponik (larutan AB Mix) ke dalam wadah bagian bawah hingga menyentuh separuh sumbu flanel, lalu letakkan di area yang terkena sinar matahari cukup."""

bahasa_indonesia_p1 = [
    {
        "number": 1,
        "materi": "Menemukan Informasi Tersurat Teks Informasi",
        "question": f"{teks_bi_p1_teks1}\n\nBerdasarkan paragraf pertama teks di atas, peran ekologis penyu sisik yang sangat penting bagi kelestarian terumbu karang adalah....",
        "options": [
            "Menghasilkan zat kapur alami yang mempercepat pembentukan karang baru",
            "Memakan spons laut agar tidak tumbuh berlebihan dan menutupi terumbu karang",
            "Menjaga kejernihan air laut dengan menyaring endapan lumpur pantai",
            "Menjadi tempat bertenggernya aneka jenis burung laut pemangsa"
        ],
        "correct": "B",
        "explanation": "Pada paragraf pertama secara tersurat dinyatakan: 'Makanan utama penyu sisik adalah spons laut (hewan porifera) yang hidup menempel pada karang. Dengan memakan spons laut tersebut, penyu sisik mengontrol populasinya agar tidak tumbuh berlebihan dan menutupi karang.' Jadi jawaban yang tepat adalah B."
    },
    {
        "number": 2,
        "materi": "Menentukan Ide Pokok Paragraf",
        "question": "Ide pokok yang terdapat pada paragraf kedua teks 'Penyu Sisik: Sang Penjaga Terumbu Karang Nusantara' adalah....",
        "options": [
            "Jenis kerajinan tangan bernilai ekonomi dari cangkang penyu sisik",
            "Penyebab pencemaran sampah plastik di wilayah perairan laut Indonesia",
            "Ancaman kepunahan populasi penyu sisik akibat aktivitas manusia dan limbah plastik",
            "Kemiripan fisik antara kantong plastik transparan dengan ubur-ubur laut"
        ],
        "correct": "C",
        "explanation": "Paragraf kedua membahas faktor-faktor yang mengancam kelangsungan hidup penyu sisik (perburuan cangkang, pencurian telur, dan sampah plastik di laut). Kalimat utamanya terletak di awal paragraf: 'Namun, keberadaan penyu sisik saat ini berada dalam bayang-bayang ancaman kepunahan yang nyata.' Jadi ide pokoknya adalah C."
    },
    {
        "number": 3,
        "materi": "Makna Kata / Istilah Teknis dalam Konteks Kalimat",
        "question": "Makna istilah 'konservasi' pada paragraf ketiga teks tersebut adalah....",
        "options": [
            "Penelitian ilmiah untuk menciptakan varietas satwa unggul",
            "Upaya pemeliharaan dan perlindungan secara teratur untuk mencegah kepunahan",
            "Pemanfaatan sumber daya alam secara maksimal demi keuntungan ekonomi",
            "Penyebaran satwa liar ke berbagai daerah baru tanpa pengawasan"
        ],
        "correct": "B",
        "explanation": "Dalam KBBI dan konteks pelestarian lingkungan, 'konservasi' bermakna upaya perlindungan, pelestarian, dan pengelolaan sumber daya alam hayati secara bijaksana untuk mencegah kepunahan. Jadi pilihan yang tepat adalah B."
    },
    {
        "number": 4,
        "materi": "Mengidentifikasi Hubungan Sebab-Akibat Berdasarkan Teks",
        "question": "Berdasarkan teks bacaan, mengapa sampah kantong plastik sangat berbahaya dan sering tertelan oleh penyu sisik di lautan?",
        "options": [
            "Kantong plastik mengeluarkan aroma amis yang memancing selera makan penyu",
            "Penyu sisik mengira kantong plastik transparan yang terapung sebagai ubur-ubur",
            "Sampah plastik menempel erat pada spons laut yang menjadi mangsa penyu",
            "Penyu sengaja memakan plastik untuk membantu proses pencernaan cangkangnya"
        ],
        "correct": "B",
        "explanation": "Pada paragraf kedua dijelaskan: 'Penyu kerap mengira kantong plastik transparan yang terapung di perairan sebagai ubur-ubur santapan mereka. Akibatnya, plastik menyumbat saluran pencernaan...' Jadi penyebab utamanya adalah kemiripan visual dengan ubur-ubur (Pilihan B)."
    },
    {
        "number": 5,
        "materi": "Menilai Kesesuaian Pernyataan Berdasarkan Teks",
        "question": "Pernyataan berikut yang TIDAK SESUAI dengan isi teks 'Penyu Sisik: Sang Penjaga Terumbu Karang Nusantara' adalah....",
        "options": [
            "Penyu sisik memiliki bentuk paruh runcing menyerupai paruh burung elang.",
            "Tukik dilepasliarkan pada siang hari terik agar tubuhnya cepat kuat berenang.",
            "Kepulauan Derawan dan Sukamade merupakan contoh lokasi konservasi penyu.",
            "Populasi penyu sisik berkurang salah satunya karena perburuan cangkang sisiknya."
        ],
        "correct": "B",
        "explanation": "Pada paragraf ketiga dinyatakan: 'Setelah menetas, anak penyu (tukik) dirawat secara intensif selama beberapa hari sebelum akhirnya dilepasliarkan kembali ke laut lepas pada malam atau dini hari guna menghindari serangan burung pemangsa.' Oleh karena itu, pernyataan B (dilepasliarkan pada siang hari terik) tidak sesuai/salah."
    },
    {
        "number": 6,
        "materi": "Menentukan Watak Tokoh dalam Cerita Anak",
        "question": f"{teks_bi_p1_teks2}\n\nWatak utama tokoh Arga yang tergambar paling dominan dalam cerita di atas adalah....",
        "options": [
            "Pemberani dan pantang menyerah",
            "Jujur, berbudi pekerti luhur, dan berbakti",
            "Pendiam dan mudah putus asa",
            "Penyabar dan gemar berpetualang"
        ],
        "correct": "B",
        "explanation": "Arga membantu ibunya menjual sayur (berbakti), dan meskipun tergoda untuk membeli sepatu baru yang rusak, ia tetap mengembalikan kantong berisi uang kepada Kakek Suro (jujur dan berbudi pekerti luhur). Jadi watak dominan Arga adalah pilihan B."
    },
    {
        "number": 7,
        "materi": "Mengidentifikasi Latar Tempat dan Suasana Cerita",
        "question": "Latar tempat terjadinya peristiwa penemuan kantong belacu cokelat oleh Arga adalah....",
        "options": [
            "Di tengah pasar desa yang ramai",
            "Di tepi semak rumput gajah dekat tikungan jembatan kayu perbatasan desa",
            "Di pekarangan rumah gubuk milik Kakek Suro",
            "Di ruang tunggu puskesmas desa"
        ],
        "correct": "B",
        "explanation": "Pada paragraf pertama tertulis jelas: 'Ketika melewati tikungan jembatan kayu di perbatasan desa, roda depan sepedanya menabrak sesuatu... Di tepi semak-semak rumput gajah, tergeletak sebuah kantong kain belacu berwarna cokelat kusam.' Jadi latar tempatnya adalah B."
    },
    {
        "number": 8,
        "materi": "Menganalisis Konflik Batin Tokoh dalam Cerita",
        "question": "Konflik batin yang sempat dialami tokoh Arga ketika melihat isi kantong kain belacu tersebut adalah....",
        "options": [
            "Kebingungan memilih apakah harus menjual sayur ke pasar atau pulang ke rumah",
            "Keinginan menggunakan uang temuan untuk membeli sepatu baru versus ajaran moral kejujuran dari ibunya",
            "Rasa takut dituduh mencuri barang berharga milik warga dusun lain",
            "Kekhawatiran sepedanya akan rusak jika memutar arah jalan"
        ],
        "correct": "B",
        "explanation": "Pada paragraf ketiga dijelaskan adanya pertentangan batin: 'Di dalam hati Arga, sempat terlintas bisikan godaan. Sepatu sekolahnya sudah jebol di bagian tumit... Namun, seketika petuah bijak sang ibu bergaung kuat... Kejujuran adalah kehormatan hidup kita.' Ini mencerminkan konflik batin pada pilihan B."
    },
    {
        "number": 9,
        "materi": "Menemukan Amanat / Pesan Moral Cerita",
        "question": "Amanat utama yang dapat dipetik dari kisah 'Kejujuran di Balik Kantong Belacu Kakek Suro' adalah....",
        "options": [
            "Simpanlah barang berharga di rumah agar tidak tercecer di jalanan",
            "Kejujuran dan integritas adalah harta paling berharga yang tidak boleh ditukar dengan kesenangan sesaat",
            "Bantulah orang tua berjualan sayur agar mendapatkan uang saku tambahan",
            "Belilah sepatu baru yang berkualitas agar tidak mudah robek di sekolah"
        ],
        "correct": "B",
        "explanation": "Amanat cerita berpusat pada nilai integritas: jangan mengambil barang yang bukan hak kita dan tetaplah jujur dalam keadaan apa pun karena kejujuran membawa ketenteraman batin dan keberkahan. Pilihan yang tepat adalah B."
    },
    {
        "number": 10,
        "materi": "Keteladanan Sikap Tokoh Cerita",
        "question": "Perilaku terpuji tokoh Arga yang paling patut diteladani oleh siswa dalam kehidupan sehari-hari adalah....",
        "options": [
            "Menabung uang saku sekolah sebanyak-banyaknya di bank",
            "Mengembalikan barang temuan yang bukan haknya secara utuh kepada pemilik aslinya",
            "Memperbaiki sepeda ontel sendiri tanpa bantuan orang dewasa",
            "Berjalan kaki saat berangkat ke sekolah demi menghemat tenaga"
        ],
        "correct": "B",
        "explanation": "Tindakan Arga yang paling bernilai keteladanan luhur adalah kesediaannya mengantarkan dan mengembalikan dompet/kantong uang temuan kepada pemiliknya tanpa pamrih (Pilihan B)."
    },
    {
        "number": 11,
        "materi": "Menentukan Topik Utama Dialog Wawancara",
        "question": f"{teks_bi_p1_teks3}\n\nTopik pokok pembicaraan dalam wawancara antara Dika dan Ibu Ratna adalah....",
        "options": [
            "Proses pembuatan kompos dari sisa tanaman hias sekolah",
            "Sosialisasi dan manfaat pelaksanaan program Bank Sampah di lingkungan sekolah",
            "Pemilihan ketua pengurus dokter cilik Adiwiyata periode baru",
            "Teknik mendaur ulang botol plastik menjadi pot bunga hias"
        ],
        "correct": "B",
        "explanation": "Seluruh percakapan membahas latar belakang, mekanisme kerja, dan dampak positif program 'Bank Sampah Berseri' di sekolah. Jadi topik utamanya adalah B."
    },
    {
        "number": 12,
        "materi": "Menemukan Informasi Rinci dalam Dialog",
        "question": "Berdasarkan dialog di atas, keuntungan langsung yang diperoleh siswa ketika menyetorkan sampah anorganik bersih ke Bank Sampah adalah....",
        "options": [
            "Mendapatkan nilai rapor tertinggi pada mata pelajaran IPA",
            "Memperoleh saldo tabungan yang kelak dapat ditukar dengan alat tulis atau bibit tanaman",
            "Dibebaskan dari tugas piket membersihkan ruang kelas selama satu semester",
            "Diberikan hadiah uang tunai langsung oleh kepala sekolah"
        ],
        "correct": "B",
        "explanation": "Ibu Ratna menjelaskan: 'Pengurus cilik bank sampah akan menimbang... mengonversikannya menjadi saldo tabungan di buku rekening siswa. Saldo tabungan ini pada akhir semester dapat ditukarkan dengan buku tulis, pensil warna, atau bibit tanaman hias.' Pilihan yang tepat adalah B."
    },
    {
        "number": 13,
        "materi": "Mengidentifikasi Dampak Positif Berdasarkan Teks",
        "question": "Dampak lingkungan yang paling nyata dirasakan sekolah setelah program Bank Sampah berjalan adalah....",
        "options": [
            "Penjualan tanaman hias sekolah meningkat pesat ke luar kota",
            "Volume timbulan sampah sekolah yang dibuang ke TPA berkurang hingga 60%",
            "Biaya operasional kebersihan sekolah menjadi sepenuhnya gratis",
            "Semua kantin sekolah dilarang menjual makanan berkemasan"
        ],
        "correct": "B",
        "explanation": "Secara eksplisit Ibu Ratna menyatakan: 'Alhamdulillah, volume timbulan sampah sekolah yang dibuang ke TPA berkurang drastis hingga 60%.' Jadi jawabannya adalah B."
    },
    {
        "number": 14,
        "materi": "Menyimpulkan Isi Dialog Percakapan",
        "question": "Simpulan yang paling tepat dari kutipan wawancara di atas adalah....",
        "options": [
            "Pengelolaan bank sampah sekolah merupakan kegiatan wajib yang harus diawasi dinas kebersihan kota.",
            "Program Bank Sampah di sekolah sangat efektif mengedukasi siswa dalam memilah sampah dan menciptakan lingkungan bersih.",
            "Sampah anorganik bernilai ekonomis jauh lebih tinggi daripada pupuk kompos organik.",
            "Siswa kelas VI harus bertanggung jawab penuh menjadi pengurus inti bank sampah sekolah."
        ],
        "correct": "B",
        "explanation": "Simpulan yang merangkum keseluruhan isi dialog adalah bahwa keberadaan Bank Sampah sekolah sukses membentuk kebiasaan peduli lingkungan pada siswa sekaligus membuat sekolah lebih asri dan ramah lingkungan (Pilihan B)."
    },
    {
        "number": 15,
        "materi": "Melengkapi Teks Prosedur Rumpang",
        "question": f"{teks_bi_p1_teks4}\n\nKalimat yang paling tepat dan runtut untuk melengkapi langkah nomor (4) pada petunjuk hidroponik tersebut adalah....",
        "options": [
            "Gunakan cat minyak untuk mewarnai bagian luar botol plastik agar kedap cahaya.",
            "Isilah bagian atas botol tersebut dengan media tanam berupa arang sekam atau rockwool secukupnya.",
            "Rendamlah biji benih sawi di dalam air hangat selama dua hari penuh.",
            "Buanglah sisa potongan botol bagian bawah ke tempat penampungan sampah."
        ],
        "correct": "B",
        "explanation": "Setelah memasang bagian botol secara terbalik dengan sumbu (langkah 3) dan sebelum menempatkan bibit bertunas (langkah 5), wadah harus diisi media tanam (seperti arang sekam, cocopeat, atau rockwool). Kalimat B adalah langkah yang logis dan tepat."
    },
    {
        "number": 16,
        "materi": "Menganalisis Fungsi Alat pada Teks Petunjuk",
        "question": "Fungsi utama kain flanel pada sistem hidroponik sumbu (wick system) berdasarkan teks prosedur di atas adalah....",
        "options": [
            "Sebagai penahan agar botol bagian atas tidak mudah goyang dan tumpah",
            "Sebagai sumbu kapiler yang menyerap larutan nutrisi dari bawah menuju akar tanaman",
            "Untuk menyaring kotoran dan lumut yang tumbuh di dalam air tandon",
            "Sebagai pengganti sinar matahari buatan bagi bibit tanaman"
        ],
        "correct": "B",
        "explanation": "Kain flanel bekerja dengan gaya kapilaritas untuk menyerap air nutrisi dari tandon bagian bawah ke atas sehingga membasahi media tanam dan diserap oleh akar bibit sawi (Pilihan B)."
    },
    {
        "number": 17,
        "materi": "Memahami Makna Kias dalam Puisi Anak",
        "question": "Perhatikan bait puisi berikut!\n\nMentari pagi menyapa ramah\nEmbun menari di ujung dedaunan\nLangkah kaki riang melangkah\nMenyambut masa depan penuh harapan\n\nMakna ungkapan 'Embun menari di ujung dedaunan' pada bait puisi di atas adalah....",
        "options": [
            "Tetesan air embun pagi yang tampak berkilau dan bergoyang terkena hembusan angin sejuk",
            "Daun-daun berguguran ke tanah karena terkena tetesan air hujan lebat",
            "Udara dingin yang menyebabkan tanaman layu di pagi hari",
            "Cahaya matahari yang terhalang oleh kabut tebal pegunungan"
        ],
        "correct": "A",
        "explanation": "Majas personifikasi 'embun menari di ujung dedaunan' menggambarkan suasana pagi hari yang indah, di mana tetesan embun berkilau terkena sinar mentari dan tampak bergoyang lembut tertiup angin sepoi-sepoi (Pilihan A)."
    },
    {
        "number": 18,
        "materi": "Menyusun Urutan Kalimat Menjadi Paragraf Padu",
        "question": "Perhatikan kalimat-kalimat petunjuk pembuatan puding berikut!\n(1) Setelah uap panasnya berkurang, masukkan puding ke dalam lemari es hingga mengeras.\n(2) Campurkan bubuk agar-agar, gula pasir, dan air kelapa muda ke dalam panci secara merata.\n(3) Rebus adonan di atas api sedang sambil terus diaduk perlahan hingga mendidih sempurna.\n(4) Tuangkan cairan puding yang telah mendidih ke dalam cetakan yang telah disiapkan.\n\nUrutan petunjuk yang tepat dan logis agar menjadi teks prosedur yang padu adalah....",
        "options": ["(2) - (3) - (4) - (1)", "(2) - (4) - (3) - (1)", "(3) - (2) - (4) - (1)", "(3) - (4) - (2) - (1)"],
        "correct": "A",
        "explanation": "Urutan memasak puding yang logis:\n- (2) Campurkan bahan ke panci.\n- (3) Rebus di atas api sambil diaduk hingga mendidih.\n- (4) Tuangkan ke cetakan.\n- (1) Masukkan ke lemari es hingga mengeras.\nUrutan yang tepat adalah (2) - (3) - (4) - (1) (Pilihan A)."
    },
    {
        "number": 19,
        "materi": "Penggunaan Kata Baku Sesuai PUEBI/EYD",
        "question": "Kalimat berikut yang seluruhnya menggunakan kata baku secara tepat sesuai Pedoman Umum Ejaan Bahasa Indonesia (PUEBI/EYD) adalah....",
        "options": [
            "Para pasien sedang antri di apotik untuk mengambil obat resep dokter.",
            "Jadwal praktik dokter spesialis anak di klinik tersebut sudah aktif sejak pagi.",
            "Rian meminta ijin kepada gurunya karena badannya sedang tidak fit.",
            "Kwalitas produk kerajinan bambu itu sangat dikagumi para turis manca negara."
        ],
        "correct": "B",
        "explanation": "Analisis kata baku:\n- Pilihan A: salah, yang baku adalah 'antre' (bukan antri) dan 'apotek' (bukan apotik).\n- Pilihan B: BENAR, 'jadwal', 'praktik', 'aktif', 'klinik' semuanya baku.\n- Pilihan C: salah, yang baku adalah 'izin' (bukan ijin).\n- Pilihan D: salah, yang baku adalah 'kualitas' (bukan kwalitas) dan 'mancanegara' (ditulis serangkai).\nJadi, kalimat yang tepat adalah B."
    },
    {
        "number": 20,
        "materi": "Penggunaan Tanda Baca dan Huruf Kapital pada Kalimat Langsung",
        "question": "Penulisan kalimat langsung berikut yang menggunakan tanda baca dan huruf kapital secara tepat adalah....",
        "options": [
            'Ibu berpesan, "Tolong rapikan buku pelajaran di mejamu, Farhan!"',
            'Ibu berpesan: "tolong rapikan buku pelajaran di mejamu, Farhan!"',
            'Ibu berpesan, "Tolong rapikan buku pelajaran di mejamu Farhan"!',
            '"Ibu berpesan, tolong rapikan buku pelajaran di mejamu, Farhan."'
        ],
        "correct": "A",
        "explanation": "Aturan penulisan kalimat langsung:\n1. Pengiring diikuti tanda koma (,).\n2. Kalimat petikan diawali huruf kapital di dalam tanda petik (\").\n3. Tanda seru (!) atau tanda titik diletakkan SEBELUM tanda petik penutup (\").\nJadi penulisan yang benar adalah: Ibu berpesan, \"Tolong rapikan buku pelajaran di mejamu, Farhan!\" (Pilihan A)."
    }
]

# ------------------------------------------
# D. BAHASA INDONESIA PAKET 2
# ------------------------------------------
teks_bi_p2_teks1 = """Bacalah teks informasi berikut dengan cermat untuk menjawab soal nomor 1 sampai 5!

PLTS Terapung Cirata: Pelopor Energi Hijau Indonesia

Pembangkit Listrik Tenaga Surya (PLTS) Terapung Cirata yang berlokasi di atas perairan Waduk Cirata, Kabupaten Purwakarta, Jawa Barat, mencatatkan tonggak sejarah baru dalam transisi energi ramah lingkungan. Fasilitas megah ini menjadi PLTS terapung terbesar di kawasan Asia Tenggara dan ketiga terbesar di dunia. Dibangun di atas permukaan air seluas kurang lebih 250 hektare, PLTS ini memiliki kapasitas pembangkitan daya mencapai 192 megawatt-peak (MWp). Energi listrik bersih yang dihasilkan dari ratusan ribu modul panel surya fotovoltaik ini dialirkan ke gardu induk PLN untuk memenuhi kebutuhan listrik puluhan ribu rumah tangga tanpa menimbulkan polusi asap.

Pembangunan pembangkit listrik tenaga surya di atas permukaan air memiliki banyak keunggulan ilmiah dibandingkan panel surya di daratan. Pertama, pemanfaatan badan air waduk tidak mengorbankan lahan pertanian produktif atau kawasan hutan daratan. Kedua, suhu permukaan air di bawah panel surya memberikan efek pendinginan alami (*cooling effect*) yang menjaga kestabilan temperatur modul, sehingga efisiensi penyerapan sinar matahari meningkat hingga beberapa persen. Ketiga, naungan deretan panel surya di atas permukaan waduk mampu mengurangi laju penguapan air (*evaporasi*) saat musim kemarau serta menekan ledakan populasi gulma air seperti eceng gondok.

Inovasi PLTS Terapung Cirata menjadi bukti nyata keseriusan Indonesia dalam menekan emisi gas rumah kaca menuju target Net Zero Emission pada tahun 2060. Dengan beralih ke sumber energi terbarukan yang melimpah dan tidak pernah habis seperti tenaga matahari, ketergantungan terhadap bahan bakar fosil batu bara yang mencemari atmosfer dapat dikurangi secara berkelanjutan demi bumi yang lebih sehat."""

teks_bi_p2_teks2 = """Bacalah cerita inspiratif berikut dengan cermat untuk menjawab soal nomor 6 sampai 10!

Robot Perahu Pembersih Parit Karya Rian

Saluran air di sepanjang pemukiman Gang Kenanga sering kali tersumbat tumpukan sampah plastik, botol bekas, dan dedaunan kering, terutama tatkala hujan deras mengguyur. Akibatnya, air parit meluap menggenangi jalan kampung dan menebarkan bau tak sedap. Rian, siswa kelas enam yang gemar mengutak-atik barang elektronik bekas di bengkel ayahnya, merasa terpanggil untuk menciptakan solusi nyata.

Memanfaatkan dinamo mobil mainan yang sudah rusak, baterai telepon genggam bekas, baling-baling plastik, dan keranjang jaring kawat, Rian mulai merakit sebuah prototipe robot perahu pembersih sampah otomatis. Selama dua pekan, ia tekun menyolder sambungan kabel dan menguji daya apungnya di bak mandi rumah. Namun, rintangan berat terjadi pada uji coba perdana di saluran parit sungguhan. Baling-baling perahunya patah menghantam batu padas, dan wadah baterainya kemasukan air kotor hingga mengeluarkan kepulan asap.

Melihat kegagalan itu, beberapa teman sempat menyuruhnya berhenti, "Sudahlah Rian, parit kotor itu bukan urusan anak sekolah seperti kita!" Namun, Rian menolak berputus asa. Dengan tersenyum tenang, ia menyahut, "Kegagalan adalah sahabat terbaik untuk belajar. Jika baling-balingnya rapuh, aku akan menggantinya dengan bilah kipas nilon yang tebal dan melapisi kotak mesin dengan lem silikon kedap air."

Keesokan sorenya, Rian kembali menurunkan perahu robot yang telah disempurnakan ke saluran air. Kali ini keajaiban dedikasi berbuah manis. Robot perahu kecil itu meluncur gesit melawan arus, menyapu sampah plastik yang terapung, lalu mengumpulkannya ke dalam keranjang jaring di buritannya. Warga kampung yang berkumpul di pinggir jalan bertepuk tangan riuh, kagum atas keuletan dan kecerdasan anak muda yang peduli pada lingkungannya tersebut."""

teks_bi_p2_teks3 = """Bacalah teks budaya nusantara berikut untuk menjawab soal nomor 11 sampai 14!

Tradisi Marosok: Bahasa Isyarat Jari dalam Perdagangan Minangkabau

Di pasar ternak tradisional masyarakat Minangkabau di Sumatera Barat, terdapat sebuah tradisi unik nan sarat nilai etika dalam jual beli hewan ternak sapi atau kerbau yang dikenal dengan istilah Marosok. Berbeda dengan pasar pada umumnya yang dipenuhi tawar-menawar harga secara lisan yang riuh, prosesi tawar-menawar dalam Marosok berlangsung dalam keheningan total tanpa terdengar sepatah kata pun mengenai nominal uang.

Dalam tradisi Marosok, pembeli dan penjual saling berjabat tangan erat, tetapi tangan mereka diselubungi di bawah sehelai kain sarung, topi caping, atau handuk kecil. Kesepakatan harga dilakukan melalui sentuhan, remasan, dan bahasa isyarat jari-jemari yang tersembunyi. Tiap jari melambangkan nilai nominal rupiah tertentu: memegang jari telunjuk melambangkan angka satu (bisa ratusan ribu atau jutaan), menekuk ruas jari melambangkan separuh nilai, sedangkan gerakan menyilangkan jari melambangkan pengurangan atau diskon harga.

Penggunaan kain penutup dalam Marosok memiliki filosofi yang mendalam, yakni menjaga privasi harga antara penjual dan pembeli. Dengan cara ini, calon pembeli lain tidak mengetahui harga penawaran ternak tersebut, sehingga kestabilan harga pasar tetap terpelihara dan persaingan tidak sehat antarpedagang dapat dicegah. Di samping itu, Marosok menghindarkan rasa malu bagi pembeli jika tawaran harganya belum cocok. Tradisi leluhur ini menjunjung tinggi kehormatan, saling menghargai, dan kerukunan bermasyarakat."""

teks_bi_p2_teks4 = """Bacalah poster edukasi kesehatan berikut untuk menjawab soal nomor 15 dan 16!

Panduan Gizi Seimbang Generasi Hebat: Gerakan 'Isi Piringku'
Kementerian Kesehatan RI mengajak keluarga Indonesia menerapkan pedoman gizi seimbang sekali makan:
- 2/3 dari setengah piring diisi MAKANAN POKOK (nasi, jagung, singkong, ubi) sebagai sumber energi utama.
- 1/3 dari setengah piring diisi LAUK-PAUK (ikan, ayam, telur, tempe, tahu) kaya protein dan mineral.
- 2/3 dari setengah piring lainnya diisi SAYUR-MAYUR segar aneka warna sumber serat dan vitamin.
- 1/3 dari setengah piring lainnya diisi BUAH-BUAHAN lokal manis kaya antioksidan.

Lengkapi piring sehatmu dengan 4 Pilar Gaya Hidup Sehat:
(1) Cuci tangan pakai sabun dengan air mengalir sebelum dan sesudah makan.
(2) Minum air putih minimal 8 gelas per hari untuk menjaga hidrasi tubuh.
(3) Lakukan aktivitas fisik atau olahraga ringan minimal 30 menit setiap hari.
(4) Pantau berat badan secara rutin setiap bulan."""

bahasa_indonesia_p2 = [
    {
        "number": 1,
        "materi": "Menentukan Gagasan Pokok Paragraf",
        "question": f"{teks_bi_p2_teks1}\n\nGagasan utama yang terdapat pada paragraf pertama teks di atas adalah....",
        "options": [
            "Proses pemasangan modul panel surya fotovoltaik di Jawa Barat",
            "Keberhasilan PLTS Terapung Cirata sebagai pelopor pembangkit listrik tenaga surya skala besar di Indonesia",
            "Penyaluran aliran listrik PLN ke seluruh rumah tangga di pulau Jawa",
            "Rencana pembangunan waduk buatan di Kabupaten Purwakarta"
        ],
        "correct": "B",
        "explanation": "Paragraf pertama memaparkan keberadaan PLTS Terapung Cirata sebagai tonggak baru transisi energi hijau di Indonesia, dengan kapasitas 192 MWp dan menjadi terbesar di Asia Tenggara. Gagasan utamanya adalah pilihan B."
    },
    {
        "number": 2,
        "materi": "Menemukan Informasi Tersurat Bacaan",
        "question": "Berdasarkan paragraf kedua, salah satu keunggulan teknis pembangunan panel surya di atas permukaan waduk adalah....",
        "options": [
            "Memerlukan biaya pemeliharaan mesin yang lebih mahal daripada panel darat",
            "Suhu air memberikan efek pendinginan alami yang meningkatkan efisiensi penyerapan sinar matahari",
            "Mampu mengubah air waduk yang kotor menjadi air mineral layak minum",
            "Mempercepat perkembangbiakan tanaman gulma eceng gondok di waduk"
        ],
        "correct": "B",
        "explanation": "Pada paragraf kedua dinyatakan: 'Kedua, suhu permukaan air di bawah panel surya memberikan efek pendinginan alami (cooling effect) yang menjaga kestabilan temperatur modul, sehingga efisiensi penyerapan sinar matahari meningkat...' Jadi jawabannya adalah B."
    },
    {
        "number": 3,
        "materi": "Memahami Makna Istilah Khusus",
        "question": "Makna istilah 'energi terbarukan' pada paragraf ketiga teks bacaan tersebut adalah....",
        "options": [
            "Sumber energi yang berasal dari pembakaran bahan tambang minyak bumi",
            "Sumber energi yang berasal dari proses alam berkelanjutan dan tidak akan pernah habis",
            "Energi yang hanya dapat dipergunakan satu kali kemudian musnah",
            "Bahan bakar fosil yang diproduksi oleh pabrik pengolahan batu bara"
        ],
        "correct": "B",
        "explanation": "Energi terbarukan (renewable energy) adalah sumber energi yang dapat diperbarui secara alami dan berkelanjutan (seperti matahari, angin, air, dan panas bumi) serta ramah lingkungan. Pilihan yang tepat adalah B."
    },
    {
        "number": 4,
        "materi": "Menyimpulkan Hubungan Sebab-Akibat",
        "question": "Mengapa keberadaan deretan panel surya terapung dapat bermanfaat bagi kelestarian air waduk saat musim kemarau?",
        "options": [
            "Karena panel surya menyedot air waduk untuk pendingin baterai",
            "Karena naungan panel surya mengurangi laju penguapan air (evaporasi) waduk",
            "Karena panel surya mendatangkan awan hujan buatan di atas perairan",
            "Karena air waduk dialirkan keluar menuju saluran irigasi sawah"
        ],
        "correct": "B",
        "explanation": "Pada paragraf kedua tertulis: 'Ketiga, naungan deretan panel surya di atas permukaan waduk mampu mengurangi laju penguapan air (evaporasi) saat musim kemarau...' Pilihan yang benar adalah B."
    },
    {
        "number": 5,
        "materi": "Membedakan Kalimat Fakta dan Opini",
        "question": "Pernyataan berikut yang merupakan FAKTA yang dapat diverifikasi dari teks 'PLTS Terapung Cirata' adalah....",
        "options": [
            "PLTS Terapung Cirata dibangun di atas permukaan waduk dengan kapasitas daya mencapai 192 MWp.",
            "Semua pembangkit listrik tenaga batu bara di Indonesia sebaiknya ditutup tahun depan.",
            "Energi matahari adalah satu-satunya sumber energi yang disukai masyarakat Jawa Barat.",
            "Pemandangan panel surya di Waduk Cirata terlihat paling indah di seluruh dunia."
        ],
        "correct": "A",
        "explanation": "Pilihan A menyajikan data faktual angka terukur: berkapasitas 192 MWp di Waduk Cirata. Pilihan B, C, dan D merupakan opini, saran subjektif, atau klaim berlebihan. Jawaban yang tepat adalah A."
    },
    {
        "number": 6,
        "materi": "Menentukan Karakter / Watak Tokoh Cerita",
        "question": f"{teks_bi_p2_teks2}\n\nSikap pantang menyerah tokoh Rian dalam cerita di atas ditunjukkan ketika....",
        "options": [
            "Menolak ajakan teman-temannya untuk bermain layang-layang di lapangan",
            "Tidak putus asa saat robot perahunya rusak pada uji coba pertama dan segera memperbaiki kelemahannya",
            "Meminta ayahnya membelikan komponen robot perahu yang baru dan mahal dari toko",
            "Membiarkan sampah di parit menumpuk hingga musim hujan berakhir"
        ],
        "correct": "B",
        "explanation": "Ketika baling-balingnya patah dan disuruh menyerah oleh teman-temannya, Rian justru menyatakan bahwa kegagalan adalah pelajaran berharga dan ia memperbaiki baling-baling serta menyegel wadah baterai dengan silikon (Pilihan B)."
    },
    {
        "number": 7,
        "materi": "Mengidentifikasi Latar Suasana Cerita",
        "question": "Suasana yang tercipta pada bagian akhir cerita saat robot buatan Rian berhasil bekerja adalah....",
        "options": [
            "Tegang dan penuh kekhawatiran warga kampung",
            "Riuh, bangga, dan penuh decak kagum warga atas keberhasilan Rian",
            "Sunyi senyap karena hujan badai yang sangat lebat",
            "Kecewa karena perahu robot tersangkut ranting pohon"
        ],
        "correct": "B",
        "explanation": "Di akhir cerita digambarkan: 'Warga kampung yang berkumpul di pinggir jalan bertepuk tangan riuh, kagum atas keuletan dan kecerdasan anak muda yang peduli pada lingkungannya tersebut.' Suasana yang tercipta adalah bangga dan kagum (Pilihan B)."
    },
    {
        "number": 8,
        "materi": "Menganalisis Penyebab Masalah (Konflik)",
        "question": "Penyebab utama rusaknya robot perahu milik Rian pada saat uji coba pertama di parit kampung adalah....",
        "options": [
            "Baterai mobil mainannya kehabisan daya listrik di tengah jalan",
            "Baling-baling perahu membentur batu padas dan wadah baterai kemasukan air kotor",
            "Keranjang jaring tersangkut pada jembatan penyeberangan warga",
            "Rian lupa menyalakan tombol sakelar daya perahu"
        ],
        "correct": "B",
        "explanation": "Pada paragraf kedua disebutkan secara jelas: 'Baling-baling perahunya patah menghantam batu padas, dan wadah baterainya kemasukan air kotor hingga mengeluarkan kepulan asap.' Jadi penyebabnya adalah B."
    },
    {
        "number": 9,
        "materi": "Menemukan Nilai Moral / Amanat Cerita",
        "question": "Pesan moral yang ingin disampaikan pengarang melalui kisah perjuangan Rian adalah....",
        "options": [
            "Masalah kebersihan lingkungan kampung sebaiknya diserahkan kepada petugas kebersihan saja",
            "Kreativitas dan kegigihan memecahkan masalah akan membuahkan hasil nyata yang bermanfaat bagi sesama",
            "Membuat barang elektronik dari bahan bekas selalu berisiko gagal dan sia-sia",
            "Jangan pernah memperlihatkan karya buatan sendiri sebelum benar-benar sempurna"
        ],
        "correct": "B",
        "explanation": "Cerita ini menanamkan pesan bahwa kepekaan terhadap lingkungan, dipadukan dengan kreativitas dan ketekunan yang pantang menyerah, akan menghasilkan karya bermanfaat bagi masyarakat (Pilihan B)."
    },
    {
        "number": 10,
        "materi": "Mengaitkan Perilaku Tokoh dengan Peribahasa",
        "question": "Peribahasa yang paling tepat untuk menggambarkan tekad dan kegigihan tokoh Rian dalam meraih keberhasilan adalah....",
        "options": [
            "Bagai air di daun talas",
            "Di mana ada kemauan, di situ ada jalan",
            "Besar pasak daripada tiang",
            "Tong kosong nyaring bunyinya"
        ],
        "correct": "B",
        "explanation": "Peribahasa 'Di mana ada kemauan, di situ ada jalan' bermakna jika seseorang bersungguh-sungguh dan memiliki tekad kuat, ia pasti akan menemukan cara untuk mengatasi rintangan dan meraih tujuan (Pilihan B)."
    },
    {
        "number": 11,
        "materi": "Menemukan Informasi Utama Budaya Tradisi",
        "question": f"{teks_bi_p2_teks3}\n\nKeunikan utama proses transaksi jual beli dalam tradisi Marosok di Minangkabau adalah....",
        "options": [
            "Transaksi dilakukan dengan menggunakan pembayaran mata uang emas kuno",
            "Tawar-menawar harga dilakukan secara hening menggunakan bahasa isyarat jari tangan di balik kain penutup",
            "Hewan ternak harus dilelang secara terbuka di hadapan tetua adat suku",
            "Pembeli dan penjual wajib menuliskan nominal harga di atas selembar daun lontar"
        ],
        "correct": "B",
        "explanation": "Teks menjelaskan secara rinci: transaksi Marosok berlangsung hening tanpa suara lisan, pembeli dan penjual berjabat tangan di bawah kain sarung/penutup dan bernegosiasi menggunakan gerakan dan sentuhan jari-jemari tangan (Pilihan B)."
    },
    {
        "number": 12,
        "materi": "Menemukan Nilai Filosofis Tradisi Budaya",
        "question": "Berdasarkan teks di atas, salah satu tujuan mulia dari penggunaan kain penutup saat tawar-menawar Marosok adalah....",
        "options": [
            "Menghangatkan tangan pembeli dan penjual dari cuaca dingin pegunungan",
            "Menjaga kerahasiaan harga agar tidak memicu persaingan tidak sehat dan menghindarkan rasa malu pembeli",
            "Menunjukkan status sosial dan kekayaan keluarga pembeli ternak",
            "Menghitung jumlah uang kertas yang dibawa oleh calon pembeli"
        ],
        "correct": "B",
        "explanation": "Paragraf ketiga memaparkan filosofinya: menjaga privasi harga agar pembeli lain tidak tahu sehingga persaingan tidak sehat dapat dicegah, serta menjaga martabat/menghindarkan rasa malu pembeli jika tawarannya belum cocok (Pilihan B)."
    },
    {
        "number": 13,
        "materi": "Menentukan Makna Kata Berdasarkan Konteks",
        "question": "Makna kata 'privasi' dalam kalimat 'menjaga privasi harga antara penjual dan pembeli' adalah....",
        "options": [
            "Kebebasan berbelanja tanpa batasan waktu",
            "Kerahasiaan atau urusan pribadi yang terlindung dari penglihatan dan pengetahuan orang luar",
            "Keuntungan finansial yang didapatkan oleh pihak pasar",
            "Kesepakatan tertulis yang disahkan di hadapan notaris hukum"
        ],
        "correct": "B",
        "explanation": "Privasi berarti kerahasiaan atau ranah pribadi yang dijaga agar tidak diketahui atau diintervensi oleh orang lain. Pilihan yang tepat adalah B."
    },
    {
        "number": 14,
        "materi": "Ringkasan Isi Teks",
        "question": "Ringkasan yang paling padu dan tepat untuk teks 'Tradisi Marosok' adalah....",
        "options": [
            "Pasar ternak di Minangkabau hanya menjual sapi dan kerbau pada hari tertentu dengan pengawasan ketat tetua adat.",
            "Tradisi Marosok adalah cara tawar-menawar hewan ternak khas Minangkabau menggunakan isyarat jari di balik kain untuk menjaga etika, privasi harga, dan keharmonisan.",
            "Penggunaan kain sarung dalam perdagangan ternak di Sumatera Barat bertujuan untuk menutupi kekurangan fisik hewan kurban.",
            "Masyarakat Minangkabau lebih memilih berdagang secara lisan karena dianggap lebih cepat dan praktis."
        ],
        "correct": "B",
        "explanation": "Ringkasan B merangkum unsur esensial teks: tradisi Marosok, cara kerja isyarat jari di balik penutup, dan nilai filosofis menjaga etika serta privasi harga. Pilihan yang tepat adalah B."
    },
    {
        "number": 15,
        "materi": "Membaca Informasi Poster Edukasi",
        "question": f"{teks_bi_p2_teks4}\n\nBerdasarkan panduan 'Isi Piringku' dalam poster edukasi di atas, kelompok makanan yang memiliki porsi seimbang sebesar 2/3 dari setengah piring adalah....",
        "options": [
            "Makanan pokok dan sayur-mayur",
            "Lauk-pauk dan buah-buahan",
            "Makanan pokok dan lauk-pauk",
            "Sayur-mayur dan buah-buahan"
        ],
        "correct": "A",
        "explanation": "Dalam poster tertulis: Makanan pokok mengambil 2/3 dari setengah piring, dan sayur-mayur juga mengambil 2/3 dari setengah piring lainnya. Keduanya berimbang besar (Pilihan A)."
    },
    {
        "number": 16,
        "materi": "Menemukan Informasi Pendukung Poster",
        "question": "Salah satu pilar pendukung gaya hidup sehat yang dianjurkan dalam pedoman 'Isi Piringku' untuk menjaga metabolisme tubuh adalah....",
        "options": [
            "Mengurangi waktu istirahat tidur malam",
            "Minum air putih minimal 8 gelas per hari dan beraktivitas fisik minimal 30 menit sehari",
            "Menghindari makan sayuran yang berdaun hijau gelap",
            "Mengonsumsi suplemen vitamin instan setiap sebelum sarapan"
        ],
        "correct": "B",
        "explanation": "Pilar (2) dan (3) menyebutkan: Minum air putih minimal 8 gelas per hari dan aktivitas fisik/olahraga ringan minimal 30 menit sehari (Pilihan B)."
    },
    {
        "number": 17,
        "materi": "Penggunaan Konjungsi Pertentangan",
        "question": "Perhatikan kalimat rumpang berikut!\n\nHujan deras mengguyur lapangan upacara sejak pagi hari, [...] para peserta gladi bersih tetap berbaris dengan tertib dan bersemangat.\n\nKata penghubung (konjungsi) yang paling tepat untuk melengkapi kalimat majemuk di atas adalah....",
        "options": ["sehingga", "karena", "namun", "bahkan"],
        "correct": "C",
        "explanation": "Kalimat tersebut menunjukkan hubungan perlawanan/pertentangan antara situasi hujan lebat dengan semangat peserta upacara yang tidak surut. Konjungsi pertentangan yang tepat adalah 'namun' (atau 'tetapi'). Pilihan yang benar adalah C."
    },
    {
        "number": 18,
        "materi": "Penggunaan Imbuhan Memper-kan / Memper-",
        "question": "Perhatikan kalimat berikut!\n\nPemerintah kota terus berupaya (luas) jalur hijau di sepanjang trotoar demi menyerap emisi kendaraan bermotor.\n\nBentuk kata berimbuhan yang tepat untuk kata dalam tanda kurung agar membentuk kalimat yang efektif adalah....",
        "options": ["memperluas", "meluaskan", "diperluas", "memperluaskan"],
        "correct": "A",
        "explanation": "Kata dasar 'luas' merupakan kata sifat. Imbuhan pembentuk kata kerja yang menyatakan 'menjadikan lebih luas' adalah 'memper-' (memperluas). Bentuk 'memperluaskan' salah karena pleonasme (tumpang tindih afiks memper-...-kan pada kata sifat). Jadi kata yang baku dan tepat adalah 'memperluas' (Pilihan A)."
    },
    {
        "number": 19,
        "materi": "Penggunaan Tanda Baca Titik Dua (:)",
        "question": "Perhatikan kalimat berikut!\n\nKoperasi sekolah menyediakan berbagai macam alat tulis: buku gambar, jangka, busur derajat, dan pensil warna.\n\nPenggunaan tanda baca titik dua (:) pada kalimat tersebut berfungsi untuk....",
        "options": [
            "Mengawali kutipan perkataan langsung tokoh",
            "Mengawali suatu pemerincian atau perincian lengkap",
            "Memisahkan anak kalimat dari induk kalimat",
            "Menutup kalimat yang masih belum selesai disusun"
        ],
        "correct": "B",
        "explanation": "Tanda titik dua (:) digunakan pada akhir suatu pernyataan lengkap yang diikuti pemerincian atau penjelasan (buku gambar, jangka, busur derajat, pensil warna). Jadi fungsinya adalah mengawali pemerincian lengkap (Pilihan B)."
    },
    {
        "number": 20,
        "materi": "Penggunaan Huruf Kapital Sesuai Ejaan Bahasa Indonesia",
        "question": "Penulisan huruf kapital berikut yang tepat dan sesuai dengan kaidah EYD adalah....",
        "options": [
            "Rombongan wisatawan itu berlayar menyeberangi Selat Bali pada hari Minggu.",
            "Danau toba merupakan Danau vulkanik terbesar di Asia Tenggara.",
            "Ibu membeli buah Pisang ambon di pasar Tradisional.",
            "Setiap hari selasa, siswa kelas Enam memakai seragam pramuka."
        ],
        "correct": "A",
        "explanation": "Analisis ejaan:\n- Pilihan A: BENAR. Nama diri geografi (Selat Bali) huruf kapital, nama hari (Minggu) huruf kapital.\n- Pilihan B: salah, harusnya Danau Toba (T kapital), 'danau' pada danau vulkanik huruf kecil.\n- Pilihan C: salah, pisang ambon adalah nama jenis (bukan nama diri), jadi p dan a huruf kecil.\n- Pilihan D: salah, selasa harus kapital (Selasa), enam huruf kecil.\nJadi kalimat yang benar adalah A."
    }
]

# ==========================================
# 2. GENERATOR DOKUMEN WORD (.DOCX)
# ==========================================

def buat_docx(paket_data, subject_name, paket_name, filename):
    doc = docx.Document()
    
    # Page Margins: 1 inch (72 pt)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
        
    # Styles
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    
    # Title Header Box
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_title.add_run("ASESMEN STANDARISASI PENDIDIKAN DAERAH / TKA SD 2026\n")
    run_sub.font.size = Pt(10)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
    
    run_main = p_title.add_run(f"TRY OUT DIGITAL INTERAKTIF: QUIZ WEB\n{subject_name.upper()} - {paket_name.upper()}\n")
    run_main.font.size = Pt(15)
    run_main.font.bold = True
    run_main.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A) # Deep Navy Blue
    
    run_desc = p_title.add_run("Disusun Berdasarkan Kisi-Kisi Kurikulum Merdeka & Standar Try Out Resmi Pemerintah")
    run_desc.font.size = Pt(9.5)
    run_desc.font.italic = True
    run_desc.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    
    # Horizontal Divider Line
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("―" * 55).font.color.rgb = RGBColor(0x93, 0xC5, 0xFD)
    
    # Petunjuk Umum
    p_petunjuk = doc.add_paragraph()
    p_petunjuk.paragraph_format.space_after = Pt(8)
    p_petunjuk.add_run("PETUNJUK UMUM:\n").font.bold = True
    p_petunjuk.add_run("1. Tulislah nama dan identitas Anda pada lembar pengerjaan yang tersedia.\n")
    p_petunjuk.add_run("2. Periksa dan bacalah soal-soal secara teliti sebelum menjawab.\n")
    p_petunjuk.add_run("3. Naskah soal ini terdiri atas 20 butir soal pilihan ganda dengan 4 opsi pilihan (A, B, C, D).\n")
    p_petunjuk.add_run("4. Pilihlah salah satu jawaban yang Anda anggap paling benar dan tepat.")
    
    # SECTION A: NASKAH SOAL
    h_sec_a = doc.add_heading(level=2)
    h_sec_a.paragraph_format.space_before = Pt(12)
    h_sec_a.paragraph_format.space_after = Pt(8)
    run_a = h_sec_a.add_run("BAGIAN A: NASKAH SOAL UJIAN (20 BUTIR)")
    run_a.font.size = Pt(12)
    run_a.font.bold = True
    run_a.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    for item in paket_data:
        # Question text
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(6)
        p_q.paragraph_format.space_after = Pt(2)
        
        # Check if question contains passage
        q_lines = item["question"].split("\n")
        run_num = p_q.add_run(f"{item['number']}. ")
        run_num.font.bold = True
        
        # If there are multiple lines (like passage + question)
        if len(q_lines) > 1:
            for l_idx, line in enumerate(q_lines):
                if l_idx == 0:
                    p_q.add_run(line + "\n")
                else:
                    p_sub = doc.add_paragraph()
                    p_sub.paragraph_format.space_before = Pt(0)
                    p_sub.paragraph_format.space_after = Pt(2)
                    if line.startswith("Bacalah") or line.startswith("Penyu") or line.startswith("PLTS") or line.startswith("Kejujuran") or line.startswith("Robot") or line.startswith("Tradisi") or line.startswith("Wawancara") or line.startswith("Panduan"):
                        run_bold_t = p_sub.add_run(line)
                        run_bold_t.font.bold = True
                        run_bold_t.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
                    else:
                        p_sub.add_run(line)
        else:
            p_q.add_run(item["question"])
            
        # Options A, B, C, D
        for opt_idx, opt_text in enumerate(item["options"]):
            opt_letter = chr(65 + opt_idx)
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Inches(0.3)
            p_opt.paragraph_format.space_before = Pt(0)
            p_opt.paragraph_format.space_after = Pt(1)
            run_opt_l = p_opt.add_run(f"{opt_letter}. ")
            run_opt_l.font.bold = True
            p_opt.add_run(opt_text)
            
    # Page Break for Answer Key & Explanations
    doc.add_page_break()
    
    # SECTION B: KUNCI JAWABAN & PEMBAHASAN LENGKAP
    h_sec_b = doc.add_heading(level=2)
    h_sec_b.paragraph_format.space_before = Pt(8)
    h_sec_b.paragraph_format.space_after = Pt(8)
    run_b = h_sec_b.add_run("BAGIAN B: KUNCI JAWABAN & PEMBAHASAN LENGKAP")
    run_b.font.size = Pt(13)
    run_b.font.bold = True
    run_b.font.color.rgb = RGBColor(0x15, 0x80, 0x3D) # Emerald Green
    
    # Table of Quick Keys
    p_t = doc.add_paragraph()
    p_t.add_run("TABEL KUNCI JAWABAN CEPAT:\n").font.bold = True
    
    table = doc.add_table(rows=5, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Light Shading Accent 1'
    
    # Fill table with 20 items (4 rows x 5 cols)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "No 1-4"
    hdr_cells[1].text = "No 5-8"
    hdr_cells[2].text = "No 9-12"
    hdr_cells[3].text = "No 13-16"
    hdr_cells[4].text = "No 17-20"
    
    for r in range(4):
        row_cells = table.rows[r+1].cells
        for c in range(5):
            idx = c * 4 + r
            if idx < len(paket_data):
                q_num = paket_data[idx]["number"]
                k_ans = paket_data[idx]["correct"]
                row_cells[c].text = f"No {q_num}: {k_ans}"
                
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # Detailed Explanations
    for item in paket_data:
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.space_before = Pt(6)
        p_exp.paragraph_format.space_after = Pt(4)
        
        run_exp_num = p_exp.add_run(f"Soal Nomor {item['number']}  |  Kunci Jawaban: {item['correct']}\n")
        run_exp_num.font.bold = True
        run_exp_num.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        
        run_mat = p_exp.add_run(f"Materi / Lingkup Kompetensi: {item['materi']}\n")
        run_mat.font.italic = True
        run_mat.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
        
        run_exp_body = p_exp.add_run(f"Pembahasan & Langkah Penyelesaian:\n{item['explanation']}\n")
        
        # Border separator
        p_line = doc.add_paragraph()
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(4)
        p_line.add_run("―" * 45).font.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)
        
    out_path = os.path.join(TARGET_DIR, filename)
    doc.save(out_path)
    print(f"[OK] Word Document berhasil dibuat: {out_path}")
    return out_path

# ==========================================
# 3. GENERATOR EXCEL STANDAR QUIZIZZ (.XLSX)
# ==========================================

def buat_quizizz_excel(paket_data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quizizz Questions"
    
    headers = [
        "Question Text",
        "Question Type",
        "Option 1",
        "Option 2",
        "Option 3",
        "Option 4",
        "Option 5",
        "Correct Answer",
        "Time in seconds",
        "Image Link",
        "Explanation"
    ]
    
    # Header styling
    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid") # Indigo
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        
    for row_idx, item in enumerate(paket_data, 2):
        # Clean question text if multi-line
        q_text = item["question"].strip()
        opts = item["options"]
        correct_letter = item["correct"].strip().upper()
        # Option index: A=1, B=2, C=3, D=4
        correct_idx = ord(correct_letter) - ord('A') + 1
        
        ws.cell(row=row_idx, column=1, value=q_text).alignment = align_left
        ws.cell(row=row_idx, column=2, value="Multiple Choice").alignment = align_center
        ws.cell(row=row_idx, column=3, value=opts[0] if len(opts) > 0 else "").alignment = align_left
        ws.cell(row=row_idx, column=4, value=opts[1] if len(opts) > 1 else "").alignment = align_left
        ws.cell(row=row_idx, column=5, value=opts[2] if len(opts) > 2 else "").alignment = align_left
        ws.cell(row=row_idx, column=6, value=opts[3] if len(opts) > 3 else "").alignment = align_left
        ws.cell(row=row_idx, column=7, value="").alignment = align_left
        ws.cell(row=row_idx, column=8, value=correct_idx).alignment = align_center
        ws.cell(row=row_idx, column=9, value=60).alignment = align_center
        ws.cell(row=row_idx, column=10, value="").alignment = align_center
        ws.cell(row=row_idx, column=11, value=item["explanation"]).alignment = align_left
        
    # Adjust column widths
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 16
    ws.column_dimensions['I'].width = 16
    ws.column_dimensions['J'].width = 15
    ws.column_dimensions['K'].width = 45
    
    out_path = os.path.join(TARGET_DIR, filename)
    wb.save(out_path)
    print(f"[OK] Quizizz Excel berhasil dibuat: {out_path}")
    return out_path

# ==========================================
# 4. GENERATOR REPORTLAB PDF (.PDF)
# ==========================================

def buat_pdf(paket_data, subject_name, paket_name, filename):
    out_path = os.path.join(TARGET_DIR, filename)
    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    style_normal = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#1F2937')
    )
    style_title = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1 # Center
    )
    style_subtitle = ParagraphStyle(
        'CustomSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#4B5563'),
        alignment=1
    )
    style_heading = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#1E3A8A'),
        spaceBefore=10,
        spaceAfter=6
    )
    style_q_bold = ParagraphStyle(
        'CustomQBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#111827'),
        spaceBefore=6,
        spaceAfter=3
    )
    style_opt = ParagraphStyle(
        'CustomOpt',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        leftIndent=15,
        textColor=colors.HexColor('#374151')
    )
    style_exp_title = ParagraphStyle(
        'CustomExpTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#15803D'),
        spaceBefore=8,
        spaceAfter=2
    )
    style_exp_body = ParagraphStyle(
        'CustomExpBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor('#374151'),
        spaceAfter=4
    )
    
    story = []
    
    # Header
    story.append(Paragraph("<b>STANDARISASI PENDIDIKAN DAERAH / TKA SD 2026</b>", style_subtitle))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"TRY OUT DIGITAL INTERAKTIF: QUIZ WEB<br/><b>{subject_name.upper()} - {paket_name.upper()}</b>", style_title))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Disusun Berdasarkan Kisi-Kisi Kurikulum Merdeka & Standar Try Out Resmi Pemerintah", style_subtitle))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#93C5FD'), spaceBefore=2, spaceAfter=8))
    
    # Petunjuk
    petunjuk_text = "<b>PETUNJUK UMUM:</b><br/>" \
                    "1. Periksa dan bacalah setiap butir soal secara saksama sebelum menjawab.<br/>" \
                    "2. Paket ini memuat 20 butir soal pilihan ganda (A, B, C, D) berbobot standar TKA.<br/>" \
                    "3. Pilihlah salah satu jawaban yang Anda anggap paling benar dan tepat."
    story.append(Paragraph(petunjuk_text, style_normal))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#E5E7EB'), spaceBefore=2, spaceAfter=8))
    
    # Section A: Soal
    story.append(Paragraph("<b>BAGIAN A: NASKAH SOAL UJIAN (20 BUTIR)</b>", style_heading))
    story.append(Spacer(1, 4))
    
    for item in paket_data:
        q_clean = item["question"].replace("\n", "<br/>")
        story.append(Paragraph(f"<b>{item['number']}.</b> {q_clean}", style_q_bold))
        for opt_idx, opt_text in enumerate(item["options"]):
            opt_letter = chr(65 + opt_idx)
            story.append(Paragraph(f"<b>{opt_letter}.</b> {opt_text}", style_opt))
        story.append(Spacer(1, 4))
        
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#15803D'), spaceBefore=8, spaceAfter=8))
    
    # Section B: Pembahasan
    story.append(Paragraph("<b>BAGIAN B: KUNCI JAWABAN & PEMBAHASAN LENGKAP</b>", style_heading))
    story.append(Spacer(1, 4))
    
    for item in paket_data:
        story.append(Paragraph(f"<b>Soal {item['number']} | Kunci: {item['correct']}</b> &nbsp;<i>({item['materi']})</i>", style_exp_title))
        exp_clean = item["explanation"].replace("\n", "<br/>")
        story.append(Paragraph(exp_clean, style_exp_body))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#E5E7EB'), spaceBefore=2, spaceAfter=4))
        
    doc.build(story)
    print(f"[OK] ReportLab PDF berhasil dibuat: {out_path}")
    return out_path

# ==========================================
# 5. EKSEKUSI PEMBUATAN SEMUA FILE
# ==========================================

print("\n=== MEMPROSES QUIZ WEB MATEMATIKA PAKET 1 ===")
buat_docx(matematika_p1, "Matematika", "Paket 1", "Quiz_Web_Matematika_Paket_1.docx")
buat_quizizz_excel(matematika_p1, "Quiz_Web_Matematika_Paket_1.xlsx")
buat_pdf(matematika_p1, "Matematika", "Paket 1", "Quiz_Web_Matematika_Paket_1.pdf")

print("\n=== MEMPROSES QUIZ WEB MATEMATIKA PAKET 2 ===")
buat_docx(matematika_p2, "Matematika", "Paket 2", "Quiz_Web_Matematika_Paket_2.docx")
buat_quizizz_excel(matematika_p2, "Quiz_Web_Matematika_Paket_2.xlsx")
buat_pdf(matematika_p2, "Matematika", "Paket 2", "Quiz_Web_Matematika_Paket_2.pdf")

print("\n=== MEMPROSES QUIZ WEB BAHASA INDONESIA PAKET 1 ===")
buat_docx(bahasa_indonesia_p1, "Bahasa Indonesia", "Paket 1", "Quiz_Web_Bahasa_Indonesia_Paket_1.docx")
buat_quizizz_excel(bahasa_indonesia_p1, "Quiz_Web_Bahasa_Indonesia_Paket_1.xlsx")
buat_pdf(bahasa_indonesia_p1, "Bahasa Indonesia", "Paket 1", "Quiz_Web_Bahasa_Indonesia_Paket_1.pdf")

print("\n=== MEMPROSES QUIZ WEB BAHASA INDONESIA PAKET 2 ===")
buat_docx(bahasa_indonesia_p2, "Bahasa Indonesia", "Paket 2", "Quiz_Web_Bahasa_Indonesia_Paket_2.docx")
buat_quizizz_excel(bahasa_indonesia_p2, "Quiz_Web_Bahasa_Indonesia_Paket_2.xlsx")
buat_pdf(bahasa_indonesia_p2, "Bahasa Indonesia", "Paket 2", "Quiz_Web_Bahasa_Indonesia_Paket_2.pdf")

# ==========================================
# 6. SIMPAN JUGA DALAM FORMAT JSON
# ==========================================
all_quiz_web_questions = []

for item in matematika_p1:
    all_quiz_web_questions.append({
        "subject": "Matematika",
        "paket": "Paket 1",
        "title": "Quiz Web Matematika Paket 1",
        "number": item["number"],
        "materi": item["materi"],
        "question": item["question"],
        "type": "Multiple Choice",
        "options": item["options"],
        "correct": item["correct"],
        "time": 60,
        "explanation": item["explanation"]
    })

for item in matematika_p2:
    all_quiz_web_questions.append({
        "subject": "Matematika",
        "paket": "Paket 2",
        "title": "Quiz Web Matematika Paket 2",
        "number": item["number"],
        "materi": item["materi"],
        "question": item["question"],
        "type": "Multiple Choice",
        "options": item["options"],
        "correct": item["correct"],
        "time": 60,
        "explanation": item["explanation"]
    })

for item in bahasa_indonesia_p1:
    all_quiz_web_questions.append({
        "subject": "Bahasa Indonesia",
        "paket": "Paket 1",
        "title": "Quiz Web Bahasa Indonesia Paket 1",
        "number": item["number"],
        "materi": item["materi"],
        "question": item["question"],
        "type": "Multiple Choice",
        "options": item["options"],
        "correct": item["correct"],
        "time": 60,
        "explanation": item["explanation"]
    })

for item in bahasa_indonesia_p2:
    all_quiz_web_questions.append({
        "subject": "Bahasa Indonesia",
        "paket": "Paket 2",
        "title": "Quiz Web Bahasa Indonesia Paket 2",
        "number": item["number"],
        "materi": item["materi"],
        "question": item["question"],
        "type": "Multiple Choice",
        "options": item["options"],
        "correct": item["correct"],
        "time": 60,
        "explanation": item["explanation"]
    })

json_path = os.path.join(TARGET_DIR, "Quiz_Web_Questions.json")
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(all_quiz_web_questions, f, ensure_ascii=False, indent=2)

print(f"\n[OK] Master JSON Quiz Web tersimpan ({len(all_quiz_web_questions)} butir): {json_path}")
print("Semua file tersimpan di folder lokal tanpa push ke GitHub.")
