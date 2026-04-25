# 🤖 JasperServer Bot — GitHub Actions

Bot otomatis yang login ke JasperServer, download laporan, dan upload ke Google Sheets.

---

## 📋 Cara Setup (Sekali saja)

### Langkah 1 — Upload ke GitHub

1. Buka **github.com** → klik **"New repository"**
2. Nama repo bebas, misal: `jasper-bot`
3. Upload 2 file ini:
   - `jasper_bot.py`
   - `.github/workflows/jasper_bot.yml`

Recovery : F4BPZA2HQXEA518Z3GY3DM14
---

### Langkah 2 — Buat Google Service Account

Ini diperlukan agar bot bisa nulis ke Google Sheets tanpa login manual.

1. Buka [console.cloud.google.com](https://console.cloud.google.com)
2. Buat project baru (atau pakai yang ada)
3. Aktifkan **Google Sheets API** dan **Google Drive API**
4. Masuk ke **IAM & Admin → Service Accounts → Create**
5. Isi nama, klik **Create and Continue → Done**
6. Klik service account yang baru dibuat → tab **Keys → Add Key → JSON**
7. Download file JSON-nya

**Convert ke Base64** (jalankan di terminal / Colab):
```bash
# Di terminal Mac/Linux:
base64 -i your-service-account.json | tr -d '\n'

# Di Windows PowerShell:
[Convert]::ToBase64String([IO.File]::ReadAllBytes("your-service-account.json"))

# Di Google Colab:
import base64
with open('your-service-account.json', 'rb') as f:
    print(base64.b64encode(f.read()).decode())
```
Salin hasilnya (string panjang).

8. **Share Google Sheet** ke email service account (ada di file JSON, field `client_email`)
   - Buka Google Sheet → Share → paste email → Editor

---

### Langkah 3 — Simpan Secrets di GitHub

Di repo GitHub, masuk ke **Settings → Secrets and variables → Actions → New repository secret**

Tambahkan secret-secret ini:

| Nama Secret | Nilai |
|---|---|
| `JASPER_USERNAME` | `muhammad.prasetyo` |
| `JASPER_PASSWORD` | `Adminhqacc12` |
| `JASPER_BASE_URL` | `http://report.tangki.id/jasperserver` |
| `GSHEET_ID` | `1DPIh2FZBAFXCaj_AbsXiMR1qSBmt4QXcp4kQ8iZXTGg` |
| `GSHEET_CREDENTIALS_B64` | *(paste string base64 dari langkah 2)* |

---

## 🚀 Cara Run

### Run Otomatis
Bot **jalan sendiri tiap Senin–Jumat jam 07:00 WIB** tanpa perlu melakukan apapun.

### Run Manual dari HP
1. Buka repo di GitHub dari browser HP
2. Tap tab **"Actions"**
3. Tap **"🤖 JasperBot Harian"** di sidebar kiri
4. Tap tombol **"Run workflow"** (kanan atas)
5. Pilih cell yang mau dijalankan (atau "all")
6. Tap **"Run workflow"** hijau

Selesai! Bot akan jalan dan hasilnya langsung masuk ke Google Sheets.

---

## 📊 Hasil

Data akan otomatis masuk ke Google Sheets:

| Cell | Tab GSheet | Laporan |
|------|-----------|---------|
| Cell 2 | `Data` | Material Transaction Summary |
| Cell 3 | `CO` | Monitor SJ Detail (CO) |
| Cell 4 | `IP` | Monitor SJ In Progress |

---

## ❓ Troubleshooting

- **Bot gagal?** → Buka tab Actions → klik run yang gagal → lihat log → download debug screenshots
- **Jadwal tidak jalan?** → GitHub kadang delay 5-15 menit untuk scheduled workflow
- **Google Sheets tidak terupdate?** → Pastikan sheet sudah di-share ke email service account

# Tesbot

trigger schedule
