# Dokumentasi Survey Lapangan CAPEX Management

## Tentang Aplikasi

Aplikasi web berbasis Streamlit untuk mengotomatisasi pembuatan dokumen rekap foto survey lapangan CAPEX. Sebelumnya proses ini dilakukan secara manual dengan mengisi template Word dan memasukkan foto satu per satu. Aplikasi ini memungkinkan staff untuk menghasilkan dokumen yang siap cetak hanya dalam beberapa menit.

---

## Fitur Utama

- Input data kegiatan melalui form (11 field informasi proyek)
- Upload foto survey lapangan hingga 10 foto (atau 8 foto jika disertai foto APD)
- Dukungan foto Penggunaan APD dengan toggle on/off
- Deteksi otomatis orientasi foto (portrait/landscape) dan penyesuaian ukuran
- Layout foto dinamis — jika jumlah foto ganjil, foto terakhir otomatis di-merge ke tengah
- Format tanggal bahasa Indonesia
- Format nilai kontrak otomatis dengan prefix "Rp"
- Preview dokumen sebelum diunduh
- Output file Word (.docx) siap pakai
- **Seluruh proses berjalan in-memory — tidak ada foto atau data yang tersimpan di server**

---

## Keamanan Data

Aplikasi ini dirancang dengan prinsip **zero data retention**:
- Foto yang diupload hanya diproses di memori selama sesi berlangsung
- Tidak ada file yang ditulis ke disk server
- Tidak ada database atau penyimpanan eksternal
- Sesi berakhir = semua data terhapus otomatis

---

## Struktur Dokumen Output

Dokumen yang dihasilkan mengikuti template resmi PAM JAYA dengan struktur:
- Header informasi proyek (Nama Kegiatan, Lokasi, Vendor, dll.)
- Tabel foto survey (2 kolom, dinamis sesuai jumlah foto)
- Seksi Penggunaan APD (opsional, muncul jika diaktifkan)

Nama file output: `Dokumentasi [Nama Kegiatan]_[Tanggal Kegiatan].docx`

---

## Teknologi

| Library | Fungsi |
|---|---|
| Streamlit | Antarmuka web |
| python-docx | Pembuatan dokumen Word |
| Pillow | Pemrosesan dan deteksi orientasi foto |
| mammoth | Preview HTML dokumen |

---

## Instalasi Lokal

```bash
git clone https://github.com/[username]/[nama-repo].git
cd [nama-repo]
pip install -r requirements.txt
streamlit run app.py
```

---

## Deploy ke Streamlit Cloud

1. Push repository ke GitHub
2. Buka [share.streamlit.io](https://share.streamlit.io)
3. Hubungkan repository GitHub
4. Set **Main file path** ke `app.py`
5. Klik **Deploy**

---

## Batasan Penggunaan

| Parameter | Nilai |
|---|---|
| Format foto yang didukung | JPG, JPEG, PNG |
| Ukuran maksimal per foto | 10 MB |
| Maksimal foto biasa | 10 (atau 8 jika foto APD diaktifkan) |
| Maksimal foto APD | 2 |

---

## Struktur Project

```
├── app.py                  # Aplikasi utama Streamlit
├── requirements.txt        # Daftar dependensi
└── utils/
    ├── doc_builder.py      # Engine pembuatan dokumen Word
    └── preview.py          # Konversi dokumen ke HTML preview
```
