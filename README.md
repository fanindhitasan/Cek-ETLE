# ETLE Checker — Bluebird

Cek massal status ETLE (tilang elektronik) untuk daftar kendaraan, lewat `etle-pmj.id`.
Hasil: satu file Excel dengan 2 sheet — **Rekap Kendaraan** (ringkasan per kendaraan)
dan **Detail Pelanggaran** (satu baris per pelanggaran, termasuk link "Detail" langsung
ke halaman aslinya kalau tersedia).

## 1. Cara termudah (Google Colab, tanpa install apa pun)

Klik badge ini lalu tinggal ikuti sel-selnya (upload file input → jalankan → download hasil):

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/fanindhitasan/Cek-ETLE/blob/main/run_etle_check.ipynb)

> Ganti `<ORG>/<REPO>` di atas dengan nama org/repo GitHub setelah repo ini dibuat (lihat langkah deploy di chat).

## 2. Cara manual (PC sendiri, Windows/PowerShell)

```powershell
git clone https://github.com/fanindhitasan/Cek-ETLE
cd <Cek-ETLE>
py -m pip install -r requirements.txt
py etle_detail_otomatis.py --input "master_kendaraan.xlsx" --sheet 0
```

### Argumen

| Argumen     | Wajib | Default              | Keterangan                                   |
|-------------|-------|-----------------------|-----------------------------------------------|
| `--input`   | Ya    | -                     | Path file Excel master kendaraan              |
| `--sheet`   | Tidak | `0`                   | Nama atau indeks sheet di file input          |
| `--output`  | Tidak | otomatis (timestamp)  | Nama file Excel hasil                         |
| `--delay`   | Tidak | `2.0` detik           | Jeda antar-request (min. 2 detik)             |
| `--timeout` | Tidak | `45` detik            | Timeout per request                           |

### Format file input

File Excel dengan minimal 3 kolom persis dengan nama berikut:

- `Nomor Polisi`
- `Nomor Rangka`
- `Nomor Mesin`

Kolom lain boleh ada, akan diabaikan.

## Catatan penting

- Script berjalan bertahap dan **menyimpan progres setiap kendaraan selesai dicek** —
  aman dihentikan (Ctrl+C) dan dilanjut nanti tanpa kehilangan hasil yang sudah didapat.
- Kolom **Detail URL** / **Link Pelanggaran Terbaru** hanya terisi kalau halaman hasil
  memang menyertakan link (biasanya untuk pelanggaran yang belum lunas/belum diproses).
- Kalau muncul status `VERIFIKASI DIPERLUKAN` (captcha) atau `RESPONS TIDAK DIKENALI`
  secara berulang, kirim contoh HTML halaman hasilnya ke pengelola script untuk
  penyesuaian parser.
