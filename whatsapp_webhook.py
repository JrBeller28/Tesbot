import os, requests, csv
from io import StringIO
from flask import Flask, request, Response
from twilio.twiml.messaging_response import MessagingResponse
# ... (import lainnya tetap sama)

# ── CONFIG BARU ──────────────────────────────────────────────────────────────
# Link GSheet (Sudah diubah ke format CSV)
GSHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSbvA_5FOxi2-nkfz8iJbptOhDfBCLM5LnTwrVLeJ4pf1hlGjSBywsTXQYYtEjuo0DY2M63wcJmc0tP/pub?gid=263347272&single=true&output=csv"

# ── FUNGSI CEK BARANG ────────────────────────────────────────────────────────
def search_inventory(query: str) -> str:
    try:
        response = requests.get(GSHEET_CSV_URL)
        response.raise_for_status()
        
        # Baca data CSV
        f = StringIO(response.text)
        reader = csv.DictReader(f)
        
        results = []
        query = query.lower()

        for row in reader:
            # Cari di semua kolom (Nama Barang, Deskripsi, dll)
            # Sesuaikan nama kolom jika kamu tahu nama kolom spesifiknya, misal: row['Nama Barang']
            combined_text = " ".join(row.values()).lower()
            
            if query in combined_text:
                # Susun tampilan per baris yang ditemukan
                item_info = [f"📦 *{v}*" if i == 0 else f"{k}: {v}" for i, (k, v) in enumerate(row.items())]
                results.append("\n".join(item_info))

        if not results:
            return f"❌ Barang *'{query}'* tidak ditemukan di katalog."

        # Batasi hasil agar chat tidak kepanjangan (misal max 5 hasil)
        header = f"🔍 *Hasil Pencarian: {query}*\n"
        body = "\n\n---\n\n".join(results[:5])
        footer = f"\n\n(Menampilkan {min(len(results), 5)} dari {len(results)} temuan)"
        
        return header + body + footer

    except Exception as e:
        return f"⚠️ Gagal mengambil data GSheet: {str(e)}"

# ── UPDATE COMMAND PARSER ────────────────────────────────────────────────────
def parse_command(text: str) -> str:
    parts = text.strip().split()
    if not parts:
        return HELP_TEXT

    cmd = parts[0].lower()

    # Perintah Status & Help tetap sama ...
    if cmd in ("help", "menu"): return HELP_TEXT
    if cmd == "status": return get_latest_run()

    # --- PERINTAH BARU: CEK ---
    if cmd == "cek":
        if len(parts) < 2:
            return "❌ Gunakan format: `cek <nama barang>`\nContoh: `cek pipa`"
        
        query = " ".join(parts[1:])
        return search_inventory(query)

    # --- PERINTAH RUN (EXISTING) ---
    if cmd == "run":
        # ... (kode run yang lama tetap di sini)
        pass

    return f"❓ Perintah tidak dikenali: `{cmd}`\nKetik `help` untuk daftar perintah."

# Update juga HELP_TEXT agar user tahu ada fitur baru
HELP_TEXT = """
🤖 *JasperBot — Menu Utama*

*cek <nama barang>*
  Cari stok barang dari Google Sheets

*status*
  Cek status JasperBot di GitHub

*run*
  Jalankan semua cell (GSheet update)

*help*
  Tampilkan menu ini
""".strip()
