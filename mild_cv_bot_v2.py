"""
MILD CV BOT v2.0 - Bot Telegram Lengkap
=========================================
✅ Database SQLite (data tidak hilang saat restart)
✅ Sistem ACTIVE/EXPIRED otomatis
✅ Pembayaran Midtrans (GoPay, QRIS, VA BCA/BNI/BRI)
✅ Menu keyboard interaktif
✅ TXT to VCF, VCF to TXT, XLSX to TXT
✅ CV Admin/Navy, Cek Duplikat
✅ Ganti Nama File, Ganti Nama Kontak
✅ Hitung Isi File, Ambil Nama File
✅ Split File, Merge File
✅ Sistem Referral + Koin
✅ Panel Admin lengkap
✅ Broadcast ke semua user
✅ Notifikasi otomatis expiry

Requirements:
    pip install python-telegram-bot==20.7 midtransclient openpyxl

Cara pakai:
    1. Ganti TOKEN dengan token dari @BotFather
    2. Ganti ADMIN_IDS dengan ID Telegram admin
    3. Isi MIDTRANS_SERVER_KEY setelah dapat dari Midtrans
    4. Jalankan: python mild_cv_bot_v2.py
"""

import io
import os
import sqlite3
import hashlib
import logging
import openpyxl

from datetime import datetime, timedelta
from contextlib import contextmanager

try:
    import midtransclient
    MIDTRANS_AVAILABLE = True
except ImportError:
    MIDTRANS_AVAILABLE = False

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    KeyboardButton,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# =====================================================
#                   KONFIGURASI
# =====================================================

TOKEN      = "8669057286:AAG7oBctKzFDqe7Axy47XFcT89LkjG7WR1U"
BOT_NAME   = "MILD CV BOT"
BOT_VER    = "v2.0"
SUPPORT    = "@MILDKE2"
ADMIN_IDS  = [7678868549]  # Ganti dengan ID Telegram admin

# ── Midtrans (isi setelah dapat key) ──────────────
MIDTRANS_SERVER_KEY    = "GANTI_MIDTRANS_SERVER_KEY"
MIDTRANS_CLIENT_KEY    = "GANTI_MIDTRANS_CLIENT_KEY"
MIDTRANS_IS_PRODUCTION = False  # Ganti True jika sudah live

# ── Harga Paket ───────────────────────────────────
PAKET = {
    "basic": {
        "nama"  : "⭐ Basic",
        "harga" : 10000,
        "durasi": 7,
        "desc"  : "Akses 7 hari semua fitur",
    },
    "pro": {
        "nama"  : "💎 Pro",
        "harga" : 25000,
        "durasi": 30,
        "desc"  : "Akses 30 hari semua fitur",
    },
    "ultimate": {
        "nama"  : "🏆 Ultimate",
        "harga" : 50000,
        "durasi": 90,
        "desc"  : "Akses 90 hari semua fitur",
    },
}

# ── Database ──────────────────────────────────────
DB_PATH = "mild_cv_bot.db"

# ── Banner ────────────────────────────────────────
BANNER = (
    "╔═══════════════════════╗\n"
    f"║  🔥  {BOT_NAME}  🔥  ║\n"
    "║   ⚡ FAST & SECURE ⚡   ║\n"
    f"║      {BOT_VER} • AUTO      ║\n"
    "╚═══════════════════════╝"
)

# =====================================================
#                   LOGGING
# =====================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# =====================================================
#                   DATABASE
# =====================================================

def init_db():
    """Buat tabel database jika belum ada."""
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        # Tabel users
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id     INTEGER PRIMARY KEY,
                nama        TEXT,
                username    TEXT,
                aktif       INTEGER DEFAULT 0,
                paket       TEXT,
                expiry      TEXT,
                koin        INTEGER DEFAULT 0,
                referral    TEXT,
                referred_by INTEGER,
                join_date   TEXT,
                file_nama   TEXT DEFAULT 'kontak'
            )
        """)
        # Tabel transaksi
        c.execute("""
            CREATE TABLE IF NOT EXISTS transaksi (
                order_id    TEXT PRIMARY KEY,
                user_id     INTEGER,
                paket       TEXT,
                jumlah      INTEGER,
                metode      TEXT,
                status      TEXT DEFAULT 'pending',
                waktu       TEXT
            )
        """)
        # Tabel kontak per user
        c.execute("""
            CREATE TABLE IF NOT EXISTS kontak (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                nama    TEXT,
                nomor   TEXT
            )
        """)
        conn.commit()


@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def get_user(user_id: int) -> dict:
    with get_db() as conn:
        row = conn.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
        if row:
            return dict(row)
        return None


def buat_user(user_id: int, nama: str, username: str):
    referral = hashlib.md5(str(user_id).encode()).hexdigest()[:8].upper()
    with get_db() as conn:
        conn.execute("""
            INSERT OR IGNORE INTO users
            (user_id, nama, username, join_date, referral)
            VALUES (?, ?, ?, ?, ?)
        """, (user_id, nama, username or "", datetime.now().strftime("%Y-%m-%d"), referral))
        conn.commit()


def update_user(user_id: int, **kwargs):
    if not kwargs:
        return
    sets  = ", ".join(f"{k}=?" for k in kwargs)
    vals  = list(kwargs.values()) + [user_id]
    with get_db() as conn:
        conn.execute(f"UPDATE users SET {sets} WHERE user_id=?", vals)
        conn.commit()


def cek_expired(user_id: int):
    u = get_user(user_id)
    if u and u["aktif"] and u["expiry"]:
        if datetime.now() > datetime.fromisoformat(u["expiry"]):
            update_user(user_id, aktif=0, paket=None, expiry=None)


def aktifkan_akses(user_id: int, paket: str):
    durasi = PAKET[paket]["durasi"]
    expiry = (datetime.now() + timedelta(days=durasi)).isoformat()
    update_user(user_id, aktif=1, paket=paket, expiry=expiry)


def sisa_waktu(user_id: int) -> str:
    u = get_user(user_id)
    if not u or not u["aktif"] or not u["expiry"]:
        return "Akses Berakhir"
    delta = datetime.fromisoformat(u["expiry"]) - datetime.now()
    if delta.total_seconds() <= 0:
        return "Akses Berakhir"
    hari  = delta.days
    jam   = delta.seconds // 3600
    menit = (delta.seconds % 3600) // 60
    if hari > 0:
        return f"{hari} Hari {jam} Jam"
    elif jam > 0:
        return f"{jam} Jam {menit} Menit"
    return f"{menit} Menit"


def simpan_transaksi(order_id, user_id, paket, jumlah, metode):
    with get_db() as conn:
        conn.execute("""
            INSERT INTO transaksi (order_id, user_id, paket, jumlah, metode, status, waktu)
            VALUES (?, ?, ?, ?, ?, 'pending', ?)
        """, (order_id, user_id, paket, jumlah, metode, datetime.now().isoformat()))
        conn.commit()


def update_transaksi(order_id, status):
    with get_db() as conn:
        conn.execute("UPDATE transaksi SET status=? WHERE order_id=?", (status, order_id))
        conn.commit()


def get_transaksi(order_id) -> dict:
    with get_db() as conn:
        row = conn.execute("SELECT * FROM transaksi WHERE order_id=?", (order_id,)).fetchone()
        return dict(row) if row else None


def get_kontak_user(user_id: int) -> list:
    with get_db() as conn:
        rows = conn.execute("SELECT nama, nomor FROM kontak WHERE user_id=?", (user_id,)).fetchall()
        return [dict(r) for r in rows]


def tambah_kontak(user_id: int, nama: str, nomor: str):
    with get_db() as conn:
        conn.execute("INSERT INTO kontak (user_id, nama, nomor) VALUES (?, ?, ?)", (user_id, nama, nomor))
        conn.commit()


def hapus_kontak(user_id: int, query: str):
    with get_db() as conn:
        conn.execute("""
            DELETE FROM kontak WHERE user_id=?
            AND (LOWER(nama) LIKE ? OR nomor LIKE ?)
        """, (user_id, f"%{query.lower()}%", f"%{query}%"))
        conn.commit()


def semua_user_ids() -> list:
    with get_db() as conn:
        rows = conn.execute("SELECT user_id FROM users").fetchall()
        return [r["user_id"] for r in rows]


# =====================================================
#                   HELPER
# =====================================================

def format_rupiah(n: int) -> str:
    return f"Rp {n:,.0f}".replace(",", ".")


def buat_order_id(user_id: int, paket: str) -> str:
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"MILD-{user_id}-{paket.upper()}-{ts}"


def parse_order_id(order_id: str) -> tuple:
    """Kembalikan (user_id, paket) dari order_id."""
    parts = order_id.split("-")
    return int(parts[1]), parts[2].lower()


def info_user_text(user, user_id: int) -> str:
    cek_expired(user_id)
    u      = get_user(user_id)
    status = "✅ ACTIVE" if u and u["aktif"] else "❌ EXPIRED"
    waktu  = sisa_waktu(user_id)
    uname  = f"@{user.username}" if user.username else "-"
    return (
        f"{BANNER}\n\n"
        f"👋 Selamat Datang di 🔀 *{BOT_NAME}*\n\n"
        f"👤 Nama: {user.full_name}\n"
        f"🆔 ID: `{user_id}`\n"
        f"🔗 Username: {uname}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"🗂 Status: {status}\n"
        f"⏰ Sisa Waktu: {waktu}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"📞 Support: {SUPPORT}\n\n"
        f"🔥 Klik fitur di bawah ini:"
    )


def main_keyboard():
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📋 TXT to VCF"),    KeyboardButton("🔄 VCF to TXT"),    KeyboardButton("📊 XLSX to TXT")],
            [KeyboardButton("📁 CV Admin/Navy"),  KeyboardButton("🔍 Cek Duplikat")],
            [KeyboardButton("✏️ Ganti Nama File"), KeyboardButton("✏️ Ganti Nama Kontak")],
            [KeyboardButton("🔢 Hitung Isi File"), KeyboardButton("📌 Ambil Nama File")],
            [KeyboardButton("✂️ Split File"),      KeyboardButton("🔀 Merge File")],
            [KeyboardButton("💳 Beli Akses"),     KeyboardButton("👤 Status Akun")],
            [KeyboardButton("🪙 Referral"),        KeyboardButton("❓ Bantuan")],
        ],
        resize_keyboard=True,
    )


def inline_testimoni():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("⭐ TESTIMONI", url=f"https://t.me/{SUPPORT.replace('@','')}"),
        InlineKeyboardButton("▶️ TUTORIAL",  url=f"https://t.me/{SUPPORT.replace('@','')}"),
    ]])


def butuh_akses(user_id: int) -> bool:
    cek_expired(user_id)
    u = get_user(user_id)
    return not (u and u["aktif"])


async def kirim_expired(update: Update):
    await update.message.reply_text(
        "❌ *Akses kamu EXPIRED!*\n\nBeli akses untuk menggunakan fitur ini.",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("💳 Beli Akses Sekarang", callback_data="buka_beli"),
        ]]),
    )


# =====================================================
#              VCF / TXT / XLSX PARSER
# =====================================================

def parse_vcf(content: str) -> list:
    result, cur = [], {}
    for line in content.splitlines():
        line = line.strip()
        if line == "BEGIN:VCARD":
            cur = {}
        elif line.startswith("FN:"):
            cur["nama"] = line[3:]
        elif line.startswith("TEL"):
            cur["nomor"] = line.split(":")[-1]
        elif line == "END:VCARD":
            if "nama" in cur and "nomor" in cur:
                result.append(cur)
            cur = {}
    return result


def parse_txt(content: str) -> list:
    result = []
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        if "|" in line:
            p = line.split("|", 1)
            result.append({"nama": p[0].strip(), "nomor": p[1].strip()})
        else:
            result.append({"nama": line, "nomor": line})
    return result


def parse_xlsx(data: bytes) -> list:
    result = []
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        vals = [str(c).strip() for c in row if c and str(c).strip() != "None"]
        if len(vals) >= 2:
            result.append({"nama": vals[0], "nomor": vals[1]})
        elif len(vals) == 1:
            result.append({"nama": vals[0], "nomor": vals[0]})
    return result


def to_vcf(kontak: list) -> str:
    lines = []
    for k in kontak:
        lines += ["BEGIN:VCARD", "VERSION:3.0",
                  f"FN:{k['nama']}", f"N:{k['nama']};;;",
                  f"TEL;TYPE=CELL:{k['nomor']}", "END:VCARD"]
    return "\n".join(lines)


def to_txt(kontak: list) -> str:
    return "\n".join(f"{k['nama']}|{k['nomor']}" for k in kontak)


# =====================================================
#              MIDTRANS
# =====================================================

def buat_transaksi(order_id: str, jumlah: int, nama: str, metode: str) -> dict:
    if not MIDTRANS_AVAILABLE:
        return {"success": False, "error": "midtransclient belum diinstall"}
    snap = midtransclient.Snap(
        is_production=MIDTRANS_IS_PRODUCTION,
        server_key=MIDTRANS_SERVER_KEY,
    )
    param = {
        "transaction_details": {"order_id": order_id, "gross_amount": jumlah},
        "customer_details"   : {"first_name": nama},
        "item_details"       : [{"id": "VIP", "price": jumlah, "quantity": 1, "name": f"{BOT_NAME} VIP"}],
    }
    metode_map = {
        "gopay"  : {"payment_type": "gopay"},
        "qris"   : {"payment_type": "qris"},
        "bca_va" : {"payment_type": "bank_transfer", "bank_transfer": {"bank": "bca"}},
        "bni_va" : {"payment_type": "bank_transfer", "bank_transfer": {"bank": "bni"}},
        "bri_va" : {"payment_type": "bank_transfer", "bank_transfer": {"bank": "bri"}},
    }
    param.update(metode_map.get(metode, {}))
    try:
        result = snap.create_transaction(param)
        return {"success": True, "url": result["redirect_url"]}
    except Exception as e:
        return {"success": False, "error": str(e)}


def cek_status(order_id: str) -> str:
    if not MIDTRANS_AVAILABLE:
        return "error"
    core = midtransclient.CoreApi(
        is_production=MIDTRANS_IS_PRODUCTION,
        server_key=MIDTRANS_SERVER_KEY,
    )
    try:
        return core.transactions.status(order_id).get("transaction_status", "unknown")
    except Exception:
        return "error"


# =====================================================
#                   HANDLERS
# =====================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    user_id = user.id
    buat_user(user_id, user.full_name, user.username)

    # Proses referral
    if context.args:
        kode = context.args[0].upper()
        u    = get_user(user_id)
        if u and not u["referred_by"]:
            with get_db() as conn:
                ref = conn.execute("SELECT user_id FROM users WHERE referral=?", (kode,)).fetchone()
                if ref and ref["user_id"] != user_id:
                    update_user(user_id, referred_by=ref["user_id"], koin=(u["koin"] or 0) + 10)
                    r = get_user(ref["user_id"])
                    update_user(ref["user_id"], koin=(r["koin"] or 0) + 20)
                    try:
                        await context.bot.send_message(
                            ref["user_id"],
                            f"🎉 Teman kamu bergabung!\nKamu dapat *20 koin*. 🪙",
                            parse_mode="Markdown",
                        )
                    except Exception:
                        pass

    await update.message.reply_text(
        info_user_text(user, user_id),
        parse_mode="Markdown",
        reply_markup=main_keyboard(),
    )
    await update.message.reply_text("⬇️", reply_markup=inline_testimoni())


# ── STATUS ────────────────────────────────────────
async def status_akun(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    user_id = user.id
    cek_expired(user_id)
    u       = get_user(user_id)
    status  = "✅ ACTIVE" if u and u["aktif"] else "❌ EXPIRED"
    waktu   = sisa_waktu(user_id)
    paket   = PAKET[u["paket"]]["nama"] if u and u["paket"] else "-"
    expiry  = ""
    if u and u["expiry"]:
        expiry = "\n📅 Hingga: *" + datetime.fromisoformat(u["expiry"]).strftime("%d/%m/%Y %H:%M") + "*"
    koin    = u["koin"] if u else 0
    total_k = len(get_kontak_user(user_id))

    await update.message.reply_text(
        f"👤 *Status Akun*\n\n"
        f"Nama: *{user.full_name}*\n"
        f"ID: `{user_id}`\n"
        f"Username: @{user.username or '-'}\n\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"🗂 Status: {status}\n"
        f"📦 Paket: {paket}\n"
        f"⏰ Sisa: *{waktu}*{expiry}\n"
        f"🪙 Koin: *{koin}*\n"
        f"📇 Kontak: *{total_k}*\n"
        f"━━━━━━━━━━━━━━━━",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("💳 Beli/Perpanjang Akses", callback_data="buka_beli"),
        ]]),
    )


# ── REFERRAL ──────────────────────────────────────
async def referral_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    user_id = user.id
    u       = get_user(user_id)
    kode    = u["referral"] if u else "-"
    koin    = u["koin"] if u else 0
    bot_me  = await context.bot.get_me()
    link    = f"https://t.me/{bot_me.username}?start={kode}"
    await update.message.reply_text(
        f"🪙 *Program Referral*\n\n"
        f"Kode kamu: `{kode}`\n"
        f"Link: `{link}`\n\n"
        f"• Kamu dapat *20 koin* per teman\n"
        f"• Teman dapat *10 koin* bonus\n\n"
        f"💼 Koin kamu: *{koin}*",
        parse_mode="Markdown",
    )


# ── BANTUAN ───────────────────────────────────────
async def bantuan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"❓ *Bantuan {BOT_NAME}*\n\n"
        f"📋 *Konversi File:*\n"
        f"• TXT to VCF — Konversi .txt ke .vcf\n"
        f"• VCF to TXT — Konversi .vcf ke .txt\n"
        f"• XLSX to TXT — Konversi Excel ke .txt\n\n"
        f"📁 *Manajemen File:*\n"
        f"• CV Admin/Navy — Format kontak jadi Admin\n"
        f"• Cek Duplikat — Hapus nomor duplikat\n"
        f"• Ganti Nama File — Ubah nama file output\n"
        f"• Ganti Nama Kontak — Ubah nama kontak\n"
        f"• Hitung Isi File — Hitung jumlah kontak\n"
        f"• Ambil Nama File — Ekstrak nama kontak\n"
        f"• Split File — Pecah file jadi beberapa bagian\n"
        f"• Merge File — Gabungkan beberapa file\n\n"
        f"👤 *Akun:*\n"
        f"• Status Akun — Cek status & sisa waktu\n"
        f"• Beli Akses — Berlangganan paket\n"
        f"• Referral — Undang teman, dapat koin\n\n"
        f"📞 Support: {SUPPORT}",
        parse_mode="Markdown",
    )


# ── BELI AKSES ────────────────────────────────────
async def beli_akses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = []
    for key, p in PAKET.items():
        keyboard.append([InlineKeyboardButton(
            f"{p['nama']} — {format_rupiah(p['harga'])} ({p['durasi']} hari)",
            callback_data=f"pilihpaket_{key}",
        )])
    text = f"💳 *Pilih Paket {BOT_NAME}*\n\n"
    for p in PAKET.values():
        text += f"{p['nama']} — *{format_rupiah(p['harga'])}*\n📝 {p['desc']}\n\n"
    await update.message.reply_text(
        text, parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def buka_beli_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = []
    for key, p in PAKET.items():
        keyboard.append([InlineKeyboardButton(
            f"{p['nama']} — {format_rupiah(p['harga'])} ({p['durasi']} hari)",
            callback_data=f"pilihpaket_{key}",
        )])
    text = f"💳 *Pilih Paket {BOT_NAME}*\n\n"
    for p in PAKET.values():
        text += f"{p['nama']} — *{format_rupiah(p['harga'])}*\n📝 {p['desc']}\n\n"
    await query.edit_message_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))


async def pilih_paket_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    paket = query.data.replace("pilihpaket_", "")
    p     = PAKET.get(paket)
    if not p:
        return
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("💚 GoPay",  callback_data=f"bayar_gopay_{paket}")],
        [InlineKeyboardButton("📱 QRIS",   callback_data=f"bayar_qris_{paket}")],
        [InlineKeyboardButton("🏦 VA BCA", callback_data=f"bayar_bca_va_{paket}")],
        [InlineKeyboardButton("🏦 VA BNI", callback_data=f"bayar_bni_va_{paket}")],
        [InlineKeyboardButton("🏦 VA BRI", callback_data=f"bayar_bri_va_{paket}")],
        [InlineKeyboardButton("« Kembali", callback_data="buka_beli")],
    ])
    await query.edit_message_text(
        f"💳 *Metode Pembayaran*\n\n"
        f"Paket: {p['nama']}\n"
        f"Harga: *{format_rupiah(p['harga'])}*\n"
        f"Durasi: *{p['durasi']} hari*\n\n"
        f"Pilih metode:",
        parse_mode="Markdown",
        reply_markup=keyboard,
    )


async def proses_bayar_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query  = update.callback_query
    await query.answer("⏳ Membuat transaksi...")
    user   = query.from_user
    data   = query.data.replace("bayar_", "")
    # Format: metode_paket (metode bisa mengandung _)
    paket  = data.rsplit("_", 1)[1]
    metode = data.rsplit("_", 1)[0]
    p      = PAKET.get(paket)
    if not p:
        return

    order_id = buat_order_id(user.id, paket)
    result   = buat_transaksi(order_id, p["harga"], user.first_name, metode)

    if result["success"]:
        simpan_transaksi(order_id, user.id, paket, p["harga"], metode)
        metode_label = {
            "gopay": "GoPay", "qris": "QRIS",
            "bca_va": "VA BCA", "bni_va": "VA BNI", "bri_va": "VA BRI",
        }.get(metode, metode)
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("💳 Bayar Sekarang", url=result["url"])],
            [InlineKeyboardButton("✅ Cek Pembayaran", callback_data=f"cekbayar_{order_id}")],
            [InlineKeyboardButton("« Kembali",         callback_data=f"pilihpaket_{paket}")],
        ])
        await query.edit_message_text(
            f"✅ *Transaksi Dibuat!*\n\n"
            f"Paket: {p['nama']}\n"
            f"Metode: *{metode_label}*\n"
            f"Total: *{format_rupiah(p['harga'])}*\n"
            f"Order ID: `{order_id}`\n\n"
            f"1️⃣ Klik *Bayar Sekarang*\n"
            f"2️⃣ Selesaikan pembayaran\n"
            f"3️⃣ Klik *Cek Pembayaran* ✅\n\n"
            f"⏳ Berlaku 1 jam",
            parse_mode="Markdown",
            reply_markup=keyboard,
        )
    else:
        # Jika Midtrans belum dikonfigurasi
        await query.edit_message_text(
            f"⚠️ *Pembayaran belum aktif*\n\n"
            f"Hubungi admin untuk aktivasi manual:\n{SUPPORT}\n\n"
            f"Paket: {p['nama']}\n"
            f"Harga: *{format_rupiah(p['harga'])}*",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("📞 Hubungi Admin", url=f"https://t.me/{SUPPORT.replace('@','')}"),
            ]]),
        )


async def cek_bayar_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query    = update.callback_query
    await query.answer("🔍 Mengecek...")
    order_id = query.data.replace("cekbayar_", "")
    trx      = get_transaksi(order_id)
    if not trx:
        await query.edit_message_text("❌ Transaksi tidak ditemukan.")
        return

    status = cek_status(order_id)
    lunas  = status in ("settlement", "capture")

    if lunas:
        update_transaksi(order_id, "paid")
        aktifkan_akses(trx["user_id"], trx["paket"])
        p      = PAKET[trx["paket"]]
        u      = get_user(trx["user_id"])
        expiry = datetime.fromisoformat(u["expiry"]).strftime("%d/%m/%Y")
        # Bonus koin referral
        if u["referred_by"]:
            ref = get_user(u["referred_by"])
            if ref:
                update_user(u["referred_by"], koin=(ref["koin"] or 0) + 5)
        await query.edit_message_text(
            f"🎉 *Pembayaran Berhasil!*\n\n"
            f"Paket *{p['nama']}* aktif!\n"
            f"Berlaku hingga: *{expiry}*\n\n"
            f"Terima kasih! Selamat menggunakan {BOT_NAME} 🔥",
            parse_mode="Markdown",
        )
    else:
        await query.edit_message_text(
            f"⏳ *Belum Terbayar*\n\nOrder: `{order_id}`\nStatus: `{status}`\n\nCoba lagi setelah bayar.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🔄 Cek Lagi", callback_data=f"cekbayar_{order_id}"),
            ]]),
        )


# =====================================================
#              FITUR KONVERSI FILE
# =====================================================

async def set_waiting(update, context, mode, pesan):
    if butuh_akses(update.effective_user.id):
        await kirim_expired(update); return
    await update.message.reply_text(pesan, parse_mode="Markdown")
    context.user_data["waiting"] = mode
    context.user_data["merge_list"] = []


async def txt_to_vcf(update, context):
    await set_waiting(update, context, "txt_vcf", "📋 Kirimkan file *.txt* untuk dikonversi ke VCF.")

async def vcf_to_txt(update, context):
    await set_waiting(update, context, "vcf_txt", "🔄 Kirimkan file *.vcf* untuk dikonversi ke TXT.")

async def xlsx_to_txt(update, context):
    await set_waiting(update, context, "xlsx_txt", "📊 Kirimkan file *.xlsx* untuk dikonversi ke TXT.")

async def cv_admin(update, context):
    await set_waiting(update, context, "cv_admin", "📁 Kirimkan file *.vcf* untuk diformat jadi CV Admin/Navy.")

async def cek_duplikat(update, context):
    await set_waiting(update, context, "nodup", "🔍 Kirimkan file *.vcf* atau *.txt* untuk dihapus duplikatnya.")

async def hitung_isi(update, context):
    await set_waiting(update, context, "hitung", "🔢 Kirimkan file *.vcf* atau *.txt* untuk dihitung isinya.")

async def ambil_nama(update, context):
    await set_waiting(update, context, "ambil_nama", "📌 Kirimkan file *.vcf* untuk diambil nama kontaknya.")

async def split_file(update, context):
    await set_waiting(update, context, "split", "✂️ Kirimkan file *.vcf* yang ingin dipecah.")

async def merge_file(update, context):
    if butuh_akses(update.effective_user.id):
        await kirim_expired(update); return
    context.user_data["waiting"]    = "merge"
    context.user_data["merge_list"] = []
    await update.message.reply_text(
        "🔀 Kirimkan file-file *.vcf* satu per satu.\nSetelah selesai ketik /merge_done",
        parse_mode="Markdown",
    )

async def merge_done(update, context):
    lst = context.user_data.get("merge_list", [])
    if len(lst) < 2:
        await update.message.reply_text("⚠️ Minimal 2 file.")
        return
    all_k = [k for sublist in lst for k in sublist]
    await update.message.reply_document(
        document=to_vcf(all_k).encode(),
        filename=f"merged_{datetime.now().strftime('%Y%m%d%H%M%S')}.vcf",
        caption=f"✅ *{len(lst)} file* digabung!\nTotal: *{len(all_k)}* kontak",
        parse_mode="Markdown",
    )
    context.user_data.pop("waiting", None)
    context.user_data["merge_list"] = []


# ── HANDLE DOKUMEN ────────────────────────────────
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    waiting = context.user_data.get("waiting")
    doc     = update.message.document
    if not doc:
        return

    file       = await context.bot.get_file(doc.file_id)
    file_bytes = await file.download_as_bytearray()
    fname      = doc.file_name or "file"
    is_vcf     = fname.lower().endswith(".vcf")
    is_xlsx    = fname.lower().endswith(".xlsx")
    content    = file_bytes.decode("utf-8", errors="ignore") if not is_xlsx else None

    if waiting == "txt_vcf":
        k = parse_txt(content) if not is_vcf else parse_vcf(content)
        await update.message.reply_document(
            document=to_vcf(k).encode(),
            filename=fname.rsplit(".", 1)[0] + ".vcf",
            caption=f"✅ Dikonversi ke VCF!\nTotal: *{len(k)}* kontak",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "vcf_txt":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        await update.message.reply_document(
            document=to_txt(k).encode(),
            filename=fname.rsplit(".", 1)[0] + ".txt",
            caption=f"✅ Dikonversi ke TXT!\nTotal: *{len(k)}* kontak",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "xlsx_txt":
        if not is_xlsx:
            await update.message.reply_text("⚠️ Kirimkan file .xlsx")
            return
        k = parse_xlsx(bytes(file_bytes))
        await update.message.reply_document(
            document=to_txt(k).encode(),
            filename=fname.rsplit(".", 1)[0] + ".txt",
            caption=f"✅ XLSX → TXT!\nTotal: *{len(k)}* kontak",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "cv_admin":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        for c in k:
            if not c["nama"].lower().startswith("admin"):
                c["nama"] = f"Admin {c['nama']}"
        await update.message.reply_document(
            document=to_vcf(k).encode(),
            filename="cv_admin.vcf",
            caption=f"✅ CV Admin/Navy selesai!\nTotal: *{len(k)}* kontak",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "nodup":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        seen, unik = set(), []
        for c in k:
            if c["nomor"] not in seen:
                seen.add(c["nomor"]); unik.append(c)
        await update.message.reply_document(
            document=to_vcf(unik).encode(),
            filename=fname.rsplit(".", 1)[0] + "_nodup.vcf",
            caption=f"🔍 *Cek Duplikat Selesai*\n\nAwal: *{len(k)}*\nDuplikat: *{len(k)-len(unik)}*\nUnik: *{len(unik)}*",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "hitung":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        await update.message.reply_text(
            f"🔢 *Hasil Hitung*\n\nFile: `{fname}`\nJumlah kontak: *{len(k)}*",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "ambil_nama":
        k     = parse_vcf(content) if is_vcf else parse_txt(content)
        names = "\n".join(f"• {c['nama']}" for c in k[:50])
        extra = f"\n_...dan {len(k)-50} lainnya_" if len(k) > 50 else ""
        await update.message.reply_text(
            f"📌 *Nama Kontak*\nFile: `{fname}`\n\n{names}{extra}",
            parse_mode="Markdown",
        )
        context.user_data.pop("waiting", None)

    elif waiting == "split":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        context.user_data["split_data"] = k
        await update.message.reply_text(
            f"✅ File diterima (*{len(k)}* kontak).\nPecah menjadi berapa bagian? (ketik angka)"
        )
        context.user_data["waiting"] = "split_count"

    elif waiting == "merge":
        k = parse_vcf(content) if is_vcf else parse_txt(content)
        context.user_data.setdefault("merge_list", []).append(k)
        idx = len(context.user_data["merge_list"])
        await update.message.reply_text(
            f"✅ File ke-{idx} diterima (*{len(k)}* kontak).\nKirim file berikutnya atau /merge_done",
            parse_mode="Markdown",
        )

    else:
        await update.message.reply_text("📎 File diterima. Pilih fitur dari menu terlebih dahulu.")


# ── HANDLE TEXT ───────────────────────────────────
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text    = update.message.text
    user_id = update.effective_user.id
    waiting = context.user_data.get("waiting")
    conv    = context.user_data.get("conv")

    # Proses split count
    if waiting == "split_count":
        try:
            n    = int(text.strip())
            data = context.user_data.get("split_data", [])
            if n < 2 or n > len(data):
                await update.message.reply_text(f"⚠️ Masukkan angka antara 2 dan {len(data)}."); return
            size = max(1, len(data) // n)
            for i in range(n):
                bag = data[i*size:(i+1)*size if i < n-1 else None]
                if not bag: continue
                await update.message.reply_document(
                    document=to_vcf(bag).encode(),
                    filename=f"bagian_{i+1}.vcf",
                    caption=f"✂️ Bagian {i+1}/{n} — *{len(bag)}* kontak",
                    parse_mode="Markdown",
                )
            context.user_data.pop("waiting", None)
            context.user_data.pop("split_data", None)
        except ValueError:
            await update.message.reply_text("⚠️ Masukkan angka yang valid.")
        return

    # Proses conversation ganti nama file
    if conv == "rename_file":
        u = get_user(user_id)
        lama = u["file_nama"] if u else "kontak"
        update_user(user_id, file_nama=text.strip())
        await update.message.reply_text(f"✅ Nama file: `{lama}.vcf` → `{text.strip()}.vcf`", parse_mode="Markdown")
        context.user_data.pop("conv", None)
        return

    if conv == "rename_ctc_lama":
        context.user_data["rename_lama"] = text.strip()
        await update.message.reply_text("✏️ Ketik *nama baru*:", parse_mode="Markdown")
        context.user_data["conv"] = "rename_ctc_baru"
        return

    if conv == "rename_ctc_baru":
        lama = context.user_data.pop("rename_lama", "")
        baru = text.strip()
        with get_db() as conn:
            conn.execute("UPDATE kontak SET nama=? WHERE user_id=? AND LOWER(nama)=?", (baru, user_id, lama.lower()))
            conn.commit()
        await update.message.reply_text(f"✅ Kontak `{lama}` → `{baru}` diubah.", parse_mode="Markdown")
        context.user_data.pop("conv", None)
        return

    # Admin broadcast/beri akses
    if waiting == "admin_broadcast" and user_id in ADMIN_IDS:
        uids = semua_user_ids()
        sent = 0
        for uid in uids:
            try:
                await context.bot.send_message(uid, f"📢 *{BOT_NAME}*\n\n{text}", parse_mode="Markdown")
                sent += 1
            except Exception:
                pass
        await update.message.reply_text(f"✅ Broadcast ke *{sent}* user.", parse_mode="Markdown")
        context.user_data.pop("waiting", None)
        return

    if waiting == "admin_beriakses" and user_id in ADMIN_IDS:
        try:
            parts = text.split()
            tid   = int(parts[0])
            paket = parts[1].lower()
            aktifkan_akses(tid, paket)
            p = PAKET[paket]
            await update.message.reply_text(f"✅ Akses *{p['nama']}* diberikan ke `{tid}`.", parse_mode="Markdown")
            try:
                await context.bot.send_message(tid, f"🎉 *Akses Aktif!*\nPaket *{p['nama']}* ({p['durasi']} hari) telah aktif!", parse_mode="Markdown")
            except Exception:
                pass
        except Exception as e:
            await update.message.reply_text(f"⚠️ Error: {e}")
        context.user_data.pop("waiting", None)
        return

    # Menu keyboard
    menu = {
        "📋 TXT to VCF"        : txt_to_vcf,
        "🔄 VCF to TXT"        : vcf_to_txt,
        "📊 XLSX to TXT"       : xlsx_to_txt,
        "📁 CV Admin/Navy"     : cv_admin,
        "🔍 Cek Duplikat"      : cek_duplikat,
        "🔢 Hitung Isi File"   : hitung_isi,
        "📌 Ambil Nama File"   : ambil_nama,
        "✂️ Split File"        : split_file,
        "🔀 Merge File"        : merge_file,
        "💳 Beli Akses"        : beli_akses,
        "👤 Status Akun"       : status_akun,
        "🪙 Referral"          : referral_cmd,
        "❓ Bantuan"           : bantuan,
    }

    if text == "✏️ Ganti Nama File":
        if butuh_akses(user_id): await kirim_expired(update); return
        await update.message.reply_text("✏️ Ketik nama file baru (tanpa ekstensi):")
        context.user_data["conv"] = "rename_file"
        return

    if text == "✏️ Ganti Nama Kontak":
        if butuh_akses(user_id): await kirim_expired(update); return
        await update.message.reply_text("✏️ Ketik *nama lama* kontak:", parse_mode="Markdown")
        context.user_data["conv"] = "rename_ctc_lama"
        return

    if text in menu:
        await menu[text](update, context)


# ── ADMIN ─────────────────────────────────────────
async def admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("❌ Akses ditolak.")
        return
    with get_db() as conn:
        total_user  = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_aktif = conn.execute("SELECT COUNT(*) FROM users WHERE aktif=1").fetchone()[0]
        total_paid  = conn.execute("SELECT COUNT(*) FROM transaksi WHERE status='paid'").fetchone()[0]
        total_rev   = conn.execute("SELECT COALESCE(SUM(jumlah),0) FROM transaksi WHERE status='paid'").fetchone()[0]

    await update.message.reply_text(
        f"🔐 *Panel Admin {BOT_NAME}*\n\n"
        f"👥 Total User: *{total_user}*\n"
        f"✅ User Aktif: *{total_aktif}*\n"
        f"💳 Transaksi Lunas: *{total_paid}*\n"
        f"💰 Total Pendapatan: *{format_rupiah(total_rev)}*",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📢 Broadcast",   callback_data="adm_broadcast"),
             InlineKeyboardButton("👑 Beri Akses",  callback_data="adm_beriakses")],
            [InlineKeyboardButton("📋 Transaksi",   callback_data="adm_transaksi"),
             InlineKeyboardButton("👥 List User",   callback_data="adm_listuser")],
        ]),
    )


async def admin_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS:
        return
    if query.data == "adm_broadcast":
        await query.edit_message_text("📢 Ketik pesan broadcast:")
        context.user_data["waiting"] = "admin_broadcast"
    elif query.data == "adm_beriakses":
        await query.edit_message_text("👑 Format: `ID_USER PAKET`\nContoh: `123456789 pro`\nPaket: basic / pro / ultimate")
        context.user_data["waiting"] = "admin_beriakses"
    elif query.data == "adm_transaksi":
        with get_db() as conn:
            rows = conn.execute("SELECT * FROM transaksi ORDER BY waktu DESC LIMIT 10").fetchall()
        lines = [f"`{r['order_id']}`\n{r['paket']} | {r['status']} | {format_rupiah(r['jumlah'])}" for r in rows]
        await query.edit_message_text(
            "📋 *10 Transaksi Terakhir*\n\n" + "\n\n".join(lines) if lines else "Belum ada transaksi.",
            parse_mode="Markdown",
        )
    elif query.data == "adm_listuser":
        with get_db() as conn:
            rows = conn.execute("SELECT user_id, nama, aktif FROM users ORDER BY join_date DESC LIMIT 10").fetchall()
        lines = [f"`{r['user_id']}` — {r['nama']} — {'✅' if r['aktif'] else '❌'}" for r in rows]
        await query.edit_message_text(
            "👥 *10 User Terbaru*\n\n" + "\n".join(lines) if lines else "Belum ada user.",
            parse_mode="Markdown",
        )


# =====================================================
#                      MAIN
# =====================================================

def main():
    # Init database
    init_db()

    app = Application.builder().token(TOKEN).build()

    # Commands
    app.add_handler(CommandHandler("start",      start))
    app.add_handler(CommandHandler("admin",      admin_cmd))
    app.add_handler(CommandHandler("merge_done", merge_done))

    # Callbacks
    app.add_handler(CallbackQueryHandler(buka_beli_cb,    pattern="^buka_beli$"))
    app.add_handler(CallbackQueryHandler(pilih_paket_cb,  pattern="^pilihpaket_"))
    app.add_handler(CallbackQueryHandler(proses_bayar_cb, pattern="^bayar_"))
    app.add_handler(CallbackQueryHandler(cek_bayar_cb,    pattern="^cekbayar_"))
    app.add_handler(CallbackQueryHandler(admin_cb,        pattern="^adm_"))

    # File & text
    app.add_handler(MessageHandler(filters.Document.ALL,            handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    print(f"🤖 {BOT_NAME} {BOT_VER} berjalan...")
    app.run_polling()


if __name__ == "__main__":
    main()
