# Data generator for Matematika Paket 6 to 10
# 20 questions per package.
# Tipe: Pilihan Ganda (14), MCMA (3: no 5, 12, 18), Kategori (3: no 7, 15, 20)

paket_matematika = {}

# ==========================================
# PAKET 6 MATEMATIKA
# ==========================================
paket_matematika[6] = [
    {
        "number": 1, "materi": "Operasi Hitung Campuran Bilangan Cacah",
        "question": "Sebuah distributor sembako memiliki stok awal 85 karung beras. Setiap karung memiliki berat 25 kg. Hari ini distributor menerima pasokan tambahan sebanyak 40 karung beras dengan berat yang sama. Jika 65 karung beras langsung didistribusikan ke pasar rakyat, berapa kilogram beras yang masih tersisa di gudang distributor?",
        "options": ["1.250 kg", "1.500 kg", "1.750 kg", "2.000 kg"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Sisa karung beras di gudang = (85 + 40) - 65 = 125 - 65 = 60 karung.\n2. Total berat beras tersisa = 60 × 25 kg = 1.500 kg.\nJadi, sisa beras di gudang adalah 1.500 kg (B)."
    },
    {
        "number": 2, "materi": "Operasi Bilangan Bulat (Suhu)",
        "question": "Suhu udara di dalam lemari pembeku (freezer) mula-mula tercatat -12°C. Saat terjadi gangguan listrik, suhu di dalam freezer naik rata-rata 2°C setiap 10 menit. Setelah listrik padam selama 50 menit, petugas menyalakan generator darurat yang mampu menurunkan suhu 3°C dalam 5 menit pertama. Berapakah suhu akhir di dalam freezer?",
        "options": ["-5°C", "-2°C", "1°C", "-8°C"],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Kenaikan suhu selama 50 menit padam = (50 : 10) × 2°C = 5 × 2°C = 10°C.\n2. Suhu setelah padam = -12°C + 10°C = -2°C.\n3. Suhu setelah generator menyala dan menurunkan 3°C = -2°C - 3°C = -5°C.\nJadi, suhu akhir freezer adalah -5°C (A)."
    },
    {
        "number": 3, "materi": "Operasi Pecahan Campuran",
        "question": "Pak Danu memiliki kebun seluas 4 1/2 hektare. Sebanyak 1 3/4 hektare ditanami kelapa sawit, 1 1/2 hektare ditanami jagung hibrida, dan sisanya ditanami aneka sayuran hortikultura. Berapa hektare luas lahan kebun yang ditanami sayuran?",
        "options": ["1 1/4 hektare", "1 1/2 hektare", "1 3/8 hektare", "1 3/4 hektare"],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Total lahan sawit dan jagung = 1 3/4 + 1 1/2 = 7/4 + 3/2 = 7/4 + 6/4 = 13/4 = 3 1/4 hektare.\n2. Luas lahan sayuran = 4 1/2 - 3 1/4 = 9/2 - 13/4 = 18/4 - 13/4 = 5/4 = 1 1/4 hektare.\nJadi, luas lahan yang ditanami sayuran adalah 1 1/4 hektare (A)."
    },
    {
        "number": 4, "materi": "Kelipatan Persekutuan Terkecil (KPK)",
        "question": "Di stasiun kota, Kereta Api Logawa berangkat setiap 45 menit sekali, sedangkan Kereta Api Sancaka berangkat setiap 60 menit sekali. Jika kedua kereta api tersebut berangkat bersamaan pertama kali pada pukul 06.15, pada pukul berapa kedua kereta api tersebut akan berangkat bersamaan untuk kedua kalinya?",
        "options": ["Pukul 08.15", "Pukul 09.15", "Pukul 09.45", "Pukul 10.15"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Menentukan KPK dari 45 dan 60:\n   Faktorisasi: 45 = 3² × 5; 60 = 2² × 3 × 5.\n   KPK = 2² × 3² × 5 = 4 × 9 × 5 = 180 menit = 3 jam.\n2. Waktu berangkat bersamaan kedua kalinya = 06.15 + 3 jam = pukul 09.15.\nJadi, jawaban yang benar adalah B (Pukul 09.15)."
    },
    {
        "number": 5, "materi": "Faktor Persekutuan Terbesar (FPB)", "tipe": "MCMA",
        "question": "Sebuah sanggar seni menerima donasi 48 set kuas lukis, 72 buku sketsa, dan 96 tabung cat akrilik. Seluruh barang donasi tersebut akan dikemas ke dalam beberapa tas paket seni dengan komposisi isi yang sama banyak tanpa ada barang tersisa. Manakah pernyataan berikut yang BENAR? (Jawaban benar lebih dari satu)",
        "options": [
            "Jumlah paket seni terbanyak yang dapat dibentuk adalah 24 tas paket",
            "Setiap tas paket seni berisi 2 set kuas lukis dan 3 buku sketsa",
            "Setiap tas paket seni berisi 5 tabung cat akrilik",
            "Total seluruh perlengkapan seni di dalam setiap tas paket berjumlah 9 barang"
        ],
        "correct": ["A", "B", "D"],
        "explanation": "Langkah analisis:\n1. Menghitung FPB dari 48, 72, dan 96:\n   48 = 2⁴ × 3; 72 = 2³ × 3²; 96 = 2⁵ × 3.\n   FPB = 2³ × 3 = 8 × 3 = 24. Jadi jumlah tas paket terbanyak adalah 24 tas (Pernyataan A BENAR).\n2. Isi per paket tas:\n   - Kuas = 48 : 24 = 2 set.\n   - Buku sketsa = 72 : 24 = 3 buku. (Pernyataan B BENAR).\n   - Cat akrilik = 96 : 24 = 4 tabung. (Pernyataan C SALAH karena menyatakan 5 tabung).\n3. Total barang per tas = 2 + 3 + 4 = 9 barang. (Pernyataan D BENAR).\nJadi, pilihan jawaban yang benar adalah A, B, dan D."
    },
    {
        "number": 6, "materi": "Perbandingan dan Skala",
        "question": "Pada sebuah denah perumahan, sebidang tanah lapangan fasilitas umum digambar dengan skala 1 : 400. Lapangan tersebut berbentuk persegi panjang dengan ukuran panjang 15 cm dan lebar 10 cm pada denah. Berapakah luas sebenarnya dari lapangan perumahan tersebut?",
        "options": ["1.200 m²", "1.800 m²", "2.400 m²", "3.600 m²"],
        "correct": "C",
        "explanation": "Langkah penyelesaian:\n1. Panjang sebenarnya = 15 cm × 400 = 6.000 cm = 60 meter.\n2. Lebar sebenarnya = 10 cm × 400 = 4.000 cm = 40 meter.\n3. Luas sebenarnya = 60 m × 40 m = 2.400 m².\nJadi, luas sebenarnya adalah 2.400 m² (C)."
    },
    {
        "number": 7, "materi": "Aritmetika Sosial (Diskon dan Pembelian)", "tipe": "Kategori",
        "question": "Menjelang tahun ajaran baru, Toko Buku Cendekia memberikan promo potongan harga khusus untuk perlengkapan sekolah:\n• Paket A (Tas Sekolah): Harga label Rp150.000, diskon 20%\n• Paket B (Sepatu Sekolah): Harga label Rp200.000, diskon 15%\n• Paket C (Buku Tulis 1 Pak): Harga label Rp50.000, diskon 10%\nBerdasarkan data harga dan diskon tersebut, tentukan kebenaran dari pernyataan-pernyataan berikut!",
        "options": [
            "Besar potongan harga (diskon) untuk sepasang sepatu sekolah pada Paket B adalah Rp30.000",
            "Harga yang harus dibayarkan pembeli untuk satu tas sekolah pada Paket A setelah diskon adalah Rp120.000",
            "Jika membeli 1 Paket B dan 1 Paket C, total uang yang harus dibayarkan adalah Rp225.000"
        ],
        "correct": {"A": "Benar", "B": "Benar", "C": "Salah"},
        "explanation": "Analisis pernyataan:\n1. Diskon Paket B = 15% × Rp200.000 = Rp30.000. (Pernyataan A: BENAR).\n2. Bayar Paket A = Rp150.000 - (20% × Rp150.000) = Rp150.000 - Rp30.000 = Rp120.000. (Pernyataan B: BENAR).\n3. Bayar Paket B = Rp200.000 - Rp30.000 = Rp170.000.\n   Bayar Paket C = Rp50.000 - (10% × Rp50.000) = Rp50.000 - Rp5.000 = Rp45.000.\n   Total B + C = Rp170.000 + Rp45.000 = Rp215.000. (Pernyataan C: SALAH, tertulis Rp225.000)."
    },
    {
        "number": 8, "materi": "Satuan Waktu dan Konversi",
        "question": "Seorang atlet renang berlatih dalam tiga sesi catatan waktu: sesi pertama 1 jam 15 menit 40 detik, sesi kedua 55 menit 35 detik, dan sesi ketiga 48 menit 25 detik. Berapa total waktu latihan atlet tersebut?",
        "options": ["2 jam 59 menit 40 detik", "3 jam 00 menit 40 detik", "3 jam 09 menit 40 detik", "3 jam 15 menit 20 detik"],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Detik = 40 + 35 + 25 = 100 detik = 1 menit 40 detik.\n2. Menit = 15 + 55 + 48 + 1 (simpanan) = 119 menit = 1 jam 59 menit.\n3. Jam = 1 + 1 (simpanan) = 2 jam.\nTotal waktu = 2 jam 59 menit 40 detik (A)."
    },
    {
        "number": 9, "materi": "Kecepatan, Jarak, dan Waktu",
        "question": "Pak Rahmat mengendarai mobil dari kota X ke kota Y yang berjarak 180 km dengan kecepatan rata-rata 60 km/jam. Di tengah perjalanan, beliau beristirahat di rest area selama 30 menit. Jika Pak Rahmat berangkat dari kota X pada pukul 07.45 WIB, pada pukul berapakah beliau tiba di kota Y?",
        "options": ["10.45 WIB", "11.15 WIB", "11.30 WIB", "11.45 WIB"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Waktu tempuh murni berkendara = 180 km : 60 km/jam = 3 jam.\n2. Waktu istirahat = 30 menit.\n3. Total durasi perjalanan = 3 jam 30 menit.\n4. Waktu tiba = 07.45 + 03.30 = 11.15 WIB.\nJadi, Pak Rahmat tiba pukul 11.15 WIB (B)."
    },
    {
        "number": 10, "materi": "Satuan Berat dan Kuantitas",
        "question": "Sebuah truk pengangkut memuat 2,5 ton jagung pipil, 18 kuintal kedelai, dan 750 kg kacang tanah. Setelah dibongkar di gudang pabrik pakan ternak sebanyak 3.850 kg, berapa kilogram sisa muatan yang masih berada di atas bak truk?",
        "options": ["1.150 kg", "1.200 kg", "1.350 kg", "1.400 kg"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Konversi ke kilogram:\n   • 2,5 ton = 2,5 × 1.000 kg = 2.500 kg.\n   • 18 kuintal = 18 × 100 kg = 1.800 kg.\n   • Kacang tanah = 750 kg.\n2. Total muatan awal = 2.500 + 1.800 + 750 = 5.050 kg.\n3. Sisa muatan setelah dibongkar = 5.050 kg - 3.850 kg = 1.200 kg.\nJadi, sisa muatan adalah 1.200 kg (B)."
    },
    {
        "number": 11, "materi": "Debit dan Volume",
        "question": "Sebuah bak penampungan air berkapasitas 1.800 liter dalam keadaan kosong dialiri air melalui dua pipa kran sekaligus. Pipa kran A memiliki debit 25 liter/menit, sedangkan pipa kran B memiliki debit 20 liter/menit. Berapa menit waktu yang diperlukan agar bak penampungan tersebut terisi penuh?",
        "options": ["35 menit", "40 menit", "45 menit", "50 menit"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Total debit kedua kran bersamaan = 25 liter/menit + 20 liter/menit = 45 liter/menit.\n2. Waktu pengisian hingga penuh = Volume : Debit = 1.800 : 45 = 40 menit.\nJadi, waktu yang dibutuhkan adalah 40 menit (B)."
    },
    {
        "number": 12, "materi": "Sifat dan Karakteristik Bangun Datar", "tipe": "MCMA",
        "question": "Perhatikan ciri-ciri bangun datar berikut:\n• Memiliki empat sisi dengan sepasang sisi sejajar yang tidak sama panjang\n• Memiliki dua sudut siku-siku\n• Kedua diagonalnya saling berpotongan tetapi tidak sama panjang dan tidak saling tegak lurus\nManakah pernyataan yang BENAR mengenai bangun datar tersebut? (Jawaban benar lebih dari satu)",
        "options": [
            "Bangun datar tersebut adalah trapesium siku-siku",
            "Jumlah seluruh sudut dalamnya adalah 360°",
            "Bangun tersebut memiliki dua sumbu simetri lipat",
            "Tinggi trapesium sama dengan panjang salah satu sisi tegaknya yang membentuk sudut 90°"
        ],
        "correct": ["A", "B", "D"],
        "explanation": "Langkah analisis:\n• Bangun dengan ciri sepasang sisi sejajar dan dua sudut siku-siku adalah trapesium siku-siku (Pernyataan A BENAR).\n• Setiap segi empat memiliki total sudut dalam 360° (Pernyataan B BENAR).\n• Trapesium siku-siku tidak memiliki sumbu simetri lipat sama sekali, sehingga simetri lipat = 0 (Pernyataan C SALAH).\n• Sisi tegak yang tegak lurus dengan sisi sejajar berfungsi sebagai tinggi trapesium (Pernyataan D BENAR).\nJadi, pernyataan yang benar adalah A, B, dan D."
    },
    {
        "number": 13, "materi": "Keliling dan Luas Bangun Datar Lingkaran",
        "question": "Sebuah taman rekreasi kota memiliki lintasan jogging berbentuk lingkaran dengan diameter 56 meter. Di sepanjang tepian luar lintasan akan ditanami pohon cemara perindang dengan jarak antar-pohon 4 meter. Berapa batang pohon cemara yang diperlukan di sekeliling lintasan tersebut? (Gunakan π = 22/7)",
        "options": ["40 pohon", "44 pohon", "48 pohon", "52 pohon"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Keliling lintasan lingkaran = π × d = 22/7 × 56 m = 22 × 8 = 176 meter.\n2. Jumlah pohon yang diperlukan = Keliling : Jarak tanam = 176 : 4 = 44 pohon.\nJadi, pohon cemara yang diperlukan adalah 44 batang pohon (B)."
    },
    {
        "number": 14, "materi": "Luas Permukaan Bangun Ruang Kubus dan Balok",
        "question": "Sebuah wadah kado berbentuk kubus tanpa tutup memiliki panjang rusuk 25 cm. Bagian luar seluruh sisi tegak dan alas kotak tersebut akan dilapisi kertas kado bermotif batik. Berapakah luas minimal kertas kado yang dibutuhkan untuk membungkus kotak tersebut?",
        "options": ["2.500 cm²", "3.125 cm²", "3.750 cm²", "4.250 cm²"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Kubus tanpa tutup terdiri dari 5 buah bidang sisi bujur sangkar (4 sisi tegak + 1 alas).\n2. Luas satu bidang sisi = s × s = 25 cm × 25 cm = 625 cm².\n3. Luas permukaan tanpa tutup = 5 × 625 cm² = 3.125 cm².\nJadi, luas kertas kado yang dibutuhkan adalah 3.125 cm² (B)."
    },
    {
        "number": 15, "materi": "Volume Bangun Ruang Tabung", "tipe": "Kategori",
        "question": "Sebuah tangki penampungan minyak berbentuk tabung memiliki jari-jari alas 70 cm dan tinggi 1,5 meter (150 cm). Tangki tersebut saat ini terisi minyak sebanyak separuh (1/2) dari kapasitas tampung totalnya. (Gunakan nilai π = 22/7). Tentukan kebenaran dari pernyataan-pernyataan berikut!",
        "options": [
            "Luas bidang alas tangki minyak tersebut adalah 15.400 cm²",
            "Kapasitas volume total tabung jika terisi penuh adalah 2.310 liter",
            "Volume minyak yang ada di dalam tangki saat ini adalah 1.255 liter"
        ],
        "correct": {"A": "Benar", "B": "Benar", "C": "Salah"},
        "explanation": "Analisis pernyataan:\n1. Luas alas tabung = π × r² = 22/7 × 70 × 70 = 22 × 10 × 70 = 15.400 cm². (Pernyataan A: BENAR).\n2. Volume total tabung = Luas alas × t = 15.400 cm² × 150 cm = 2.310.000 cm³ = 2.310 liter (karena 1 liter = 1.000 cm³). (Pernyataan B: BENAR).\n3. Volume minyak separuh = 1/2 × 2.310 liter = 1.155 liter. (Pernyataan C: SALAH, karena tertulis 1.255 liter)."
    },
    {
        "number": 16, "materi": "Pengolahan Data (Diagram Batang)",
        "question": "Data penjualan seragam olahraga di koperasi sekolah selama lima hari tercatat sebagai berikut: Senin terjual 24 stel, Selasa 30 stel, Rabu 18 stel, Kamis 26 stel, dan Jumat 32 stel. Berapakah rata-rata (mean) penjualan seragam olahraga per hari di koperasi tersebut?",
        "options": ["24 stel", "26 stel", "28 stel", "30 stel"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Jumlah seluruh penjualan = 24 + 30 + 18 + 26 + 32 = 130 stel.\n2. Banyak hari penjualan = 5 hari.\n3. Rata-rata per hari = 130 : 5 = 26 stel.\nJadi, rata-rata penjualan seragam adalah 26 stel per hari (B)."
    },
    {
        "number": 17, "materi": "Statistika (Modus dan Median)",
        "question": "Nilai ulangan harian matematika dari 12 siswa kelas VI adalah sebagai berikut:\n75, 80, 85, 70, 80, 90, 85, 80, 75, 80, 95, 70.\nTentukan nilai modus dan median dari data ulangan matematika tersebut berturut-turut!",
        "options": ["Modus = 80 dan Median = 80", "Modus = 80 dan Median = 82,5", "Modus = 85 dan Median = 80", "Modus = 80 dan Median = 77,5"],
        "correct": "A",
        "explanation": "Langkah penyelesaian:\n1. Frekuensi nilai:\n   • Nilai 70 muncul 2 kali\n   • Nilai 75 muncul 2 kali\n   • Nilai 80 muncul 4 kali (Paling sering muncul, sehingga Modus = 80)\n   • Nilai 85 muncul 2 kali\n   • Nilai 90 muncul 1 kali\n   • Nilai 95 muncul 1 kali\n2. Urutan data dari terkecil ke terbesar (12 data):\n   70, 70, 75, 75, 80, [80, 80], 85, 85, 90, 95\n   Median terletak di antara data ke-6 dan ke-7: (80 + 80) : 2 = 80.\nJadi, Modus = 80 dan Median = 80 (A)."
    },
    {
        "number": 18, "materi": "Penyajian Data (Diagram Lingkaran)", "tipe": "MCMA",
        "question": "Sebuah sekolah dasar memiliki 180 siswa kelas VI yang mengikuti kegiatan ekstrakurikuler wajib dengan rincian persentase:\n• Pramuka = 40%\n• Dokter Kecil (PMR) = 25%\n• Robotik = 15%\n• Seni Tari = 20%\nManakah pernyataan berikut yang BENAR mengenai data ekstrakurikuler tersebut? (Jawaban benar lebih dari satu)",
        "options": [
            "Jumlah siswa yang mengikuti ekstrakurikuler Pramuka adalah 72 anak",
            "Jumlah siswa yang mengikuti Dokter Kecil adalah 45 anak",
            "Selisih jumlah siswa peminat Pramuka dan Robotik adalah 40 anak",
            "Peminat Seni Tari berjumlah 36 anak"
        ],
        "correct": ["A", "B", "D"],
        "explanation": "Langkah analisis:\n• Pramuka = 40% × 180 = 72 anak. (Pernyataan A BENAR).\n• Dokter Kecil = 25% × 180 = 45 anak. (Pernyataan B BENAR).\n• Robotik = 15% × 180 = 27 anak.\n  Selisih Pramuka dan Robotik = 72 - 27 = 45 anak (bukan 40 anak, jadi C SALAH).\n• Seni Tari = 20% × 180 = 36 anak. (Pernyataan D BENAR).\nJadi, pernyataan yang benar adalah A, B, dan D."
    },
    {
        "number": 19, "materi": "Koordinat Kartesius dan Bangun Datar",
        "question": "Pada bidang koordinat Kartesius, terdapat tiga titik sudut segitiga siku-siku yaitu titik K(-2, 1), titik L(4, 1), dan titik M(-2, 6). Jika satuan koordinat dihitung dalam sentimeter, berapakah luas segitiga KLM tersebut?",
        "options": ["12 cm²", "15 cm²", "20 cm²", "30 cm²"],
        "correct": "B",
        "explanation": "Langkah penyelesaian:\n1. Panjang alas KL (horizontal) = xL - xK = 4 - (-2) = 6 cm.\n2. Panjang tinggi KM (vertikal) = yM - yK = 6 - 1 = 5 cm.\n3. Luas segitiga KLM = 1/2 × alas × tinggi = 1/2 × 6 × 5 = 15 cm².\nJadi, luas segitiga adalah 15 cm² (B)."
    },
    {
        "number": 20, "materi": "Peluang Sederhana dan Peluang Empiris", "tipe": "Kategori",
        "question": "Dalam sebuah kotak undian perayaan HUT Kemerdekaan RI di sekolah, terdapat 50 kupon bernomor 1 sampai dengan 50. Seorang siswa mengambil satu lembar kupon secara acak. Tentukan kebenaran dari pernyataan-pernyataan peluang berikut!",
        "options": [
            "Peluang terambilnya kupon bernomor bilangan kelipatan 10 adalah 1/10",
            "Peluang terambilnya kupon bernomor bilangan genap adalah 1/2",
            "Jumlah kupon yang bernomor bilangan prima antara 1 sampai 50 adalah 20 lembar"
        ],
        "correct": {"A": "Benar", "B": "Benar", "C": "Salah"},
        "explanation": "Analisis pernyataan:\n1. Bilangan kelipatan 10 dari 1 sampai 50: {10, 20, 30, 40, 50} ada 5 buah. Peluang = 5/50 = 1/10. (Pernyataan A: BENAR).\n2. Bilangan genap dari 1 sampai 50 ada 25 buah. Peluang = 25/50 = 1/2. (Pernyataan B: BENAR).\n3. Bilangan prima dari 1 sampai 50 adalah {2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47} berjumlah tepat 15 buah, bukan 20 buah. (Pernyataan C: SALAH)."
    }
]

print("Paket 6 data ready!")
