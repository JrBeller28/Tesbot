#!/usr/bin/env python3
"""
WhatsApp → Twilio → GitHub Actions trigger untuk JasperBot
"""
import os, requests
from flask import Flask, request, Response
from twilio.twiml.messaging_response import MessagingResponse
from twilio.request_validator import RequestValidator

app = Flask(__name__)

# ── ENV ──────────────────────────────────────────────────────────────────────
GITHUB_TOKEN    = os.environ["GITHUB_TOKEN"]       # PAT: scope actions:write
GITHUB_OWNER    = os.environ["GITHUB_OWNER"]       # username / org
GITHUB_REPO     = os.environ["GITHUB_REPO"]        # nama repo
GITHUB_REF      = os.environ.get("GITHUB_REF", "main")
GITHUB_WF_FILE  = os.environ.get("GITHUB_WF_FILE", "jasperbot.yml")  # nama file workflow
ALLOWED_NUMBERS = set(os.environ["ALLOWED_NUMBERS"].split(","))       # +628xxx,+628yyy
TWILIO_TOKEN    = os.environ.get("TWILIO_AUTH_TOKEN", "")            # opsional, untuk validasi

# ── CELL INFO ────────────────────────────────────────────────────────────────
CELL_INFO = {
    2: "Material Transaction Summary  → tab *Data*",
    3: "Inventory Move In Progress    → tab *IM_IP*",
    4: "Monitor SJ In Progress        → tab *IP*",
    5: "iDempiere ERP                 → tab *IP_iDempiere*",
}
ALL_CELLS = [2, 3, 4, 5]

HELP_TEXT = """
🤖 *JasperBot — Perintah WhatsApp*

*run*
  Jalankan semua cell (2, 3, 4, 5)

*run 2 4*
  Jalankan cell tertentu (pisah spasi)
  Cell yang tersedia:
  • 2 → Material Transaction Summary
  • 3 → Inventory Move In Progress
  • 4 → Monitor SJ In Progress
  • 5 → iDempiere ERP

*run --deadline 14:30*
  Semua cell, selesai paling lambat 14:30 WIB

*run 2 3 --deadline 13:00*
  Cell 2 & 3, dengan deadline

*status*
  Cek status run terakhir di GitHub Actions

*help*
  Tampilkan menu ini
""".strip()


# ── GITHUB API ───────────────────────────────────────────────────────────────
def _gh_headers():
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

def trigger_workflow(cells: list[int], deadline: str = "") -> dict:
    cell_str = " ".join(str(c) for c in cells)
    url = (f"https://api.github.com/repos/{GITHUB_OWNER}/"
           f"{GITHUB_REPO}/actions/workflows/{GITHUB_WF_FILE}/dispatches")
    r = requests.post(url, headers=_gh_headers(), json={
        "ref": GITHUB_REF,
        "inputs": {"cell": cell_str, "deadline": deadline},
    })
    return {"status": r.status_code, "ok": r.status_code == 204}

def get_latest_run() -> str:
    url = (f"https://api.github.com/repos/{GITHUB_OWNER}/"
           f"{GITHUB_REPO}/actions/runs?per_page=1&event=workflow_dispatch")
    r = requests.get(url, headers=_gh_headers())
    data = r.json()
    runs = data.get("workflow_runs", [])
    if not runs:
        return "⚠️ Belum ada run yang tercatat."
    run = runs[0]
    status    = run.get("status", "-")
    conclusion = run.get("conclusion") or "-"
    emoji = {
        "completed/success":   "✅",
        "completed/failure":   "❌",
        "completed/cancelled": "🚫",
        "in_progress/-":       "🔄",
        "queued/-":            "⏳",
    }.get(f"{status}/{conclusion}", "❓")
    return (
        f"{emoji} *Run #{run['run_number']}*\n"
        f"Status    : {status}\n"
        f"Kesimpulan: {conclusion}\n"
        f"Mulai     : {run['created_at']}\n"
        f"🔗 {run['html_url']}"
    )


# ── COMMAND PARSER ───────────────────────────────────────────────────────────
def parse_command(text: str) -> str:
    parts = text.strip().split()
    if not parts:
        return HELP_TEXT

    cmd = parts[0].lower()

    # help
    if cmd in ("help", "bantuan", "menu", "?"):
        return HELP_TEXT

    # status
    if cmd in ("status", "cek", "check"):
        return get_latest_run()

    # run
    if cmd == "run":
        cells   = []
        deadline = ""
        i = 1
        while i < len(parts):
            if parts[i] == "--deadline" and i + 1 < len(parts):
                deadline = parts[i + 1]; i += 2
            else:
                try:
                    c = int(parts[i])
                    if c in CELL_INFO:
                        cells.append(c)
                    else:
                        return f"❌ Cell *{c}* tidak valid. Tersedia: 2 3 4 5"
                except ValueError:
                    return f"❌ '{parts[i]}' bukan angka cell yang valid."
                i += 1

        if not cells:
            cells = ALL_CELLS

        cells = sorted(set(cells))

        # Validasi format deadline
        if deadline:
            try:
                h, m = map(int, deadline.split(":"))
                assert 0 <= h <= 23 and 0 <= m <= 59
            except:
                return f"❌ Format deadline salah: *{deadline}*\nContoh: `--deadline 14:30`"

        result = trigger_workflow(cells=cells, deadline=deadline)
        if not result["ok"]:
            return (f"❌ Gagal trigger GitHub Actions (HTTP {result['status']}).\n"
                    f"Cek GITHUB_TOKEN / nama workflow.")

        lines = [f"✅ *JasperBot dijalankan!*\n"]
        lines.append("📋 *Cell yang akan berjalan:*")
        for c in cells:
            lines.append(f"  • Cell {c} — {CELL_INFO[c]}")
        if deadline:
            lines.append(f"\n🕐 Deadline: *{deadline} WIB*")
        lines.append("\n⏳ Ketik `status` untuk cek progres.")
        return "\n".join(lines)

    return f"❓ Perintah tidak dikenali: `{cmd}`\nKetik `help` untuk daftar perintah."


# ── WEBHOOK ──────────────────────────────────────────────────────────────────
@app.route("/webhook/whatsapp", methods=["POST"])
def whatsapp_webhook():
    # Validasi signature Twilio (aktifkan jika TWILIO_TOKEN diset)
    if TWILIO_TOKEN:
        validator = RequestValidator(TWILIO_TOKEN)
        url       = request.url
        params    = request.form.to_dict()
        sig       = request.headers.get("X-Twilio-Signature", "")
        if not validator.validate(url, params, sig):
            return Response("Forbidden", status=403)

    sender = request.form.get("From", "")
    body   = request.form.get("Body", "").strip()

    resp = MessagingResponse()
    msg  = resp.message()

    if sender not in ALLOWED_NUMBERS:
        msg.body("⛔ Nomor tidak diizinkan.")
        return Response(str(resp), mimetype="text/xml")

    print(f"[WA] {sender}: {body!r}")
    reply = parse_command(body)
    msg.body(reply)
    return Response(str(resp), mimetype="text/xml")


@app.route("/health")
def health():
    return {"status": "ok", "repo": f"{GITHUB_OWNER}/{GITHUB_REPO}"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
