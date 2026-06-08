import os, json, base64, re, io
from datetime import datetime
import httpx
import openpyxl
import psycopg2
from psycopg2.extras import Json
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]
DATABASE_URL  = os.environ.get("DATABASE_URL", "")

# ── Database ──────────────────────────────────────────────────────────────────
def get_conn():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS vedomosti (
                    id SERIAL PRIMARY KEY,
                    user_id BIGINT NOT NULL,
                    vedomist_type TEXT,
                    dilinitsa TEXT,
                    period TEXT,
                    source_file TEXT,
                    rows JSONB,
                    created_at TIMESTAMP DEFAULT NOW(),
                    UNIQUE(user_id, vedomist_type, dilinitsa, period)
                );
                CREATE TABLE IF NOT EXISTS phone_books (
                    user_id BIGINT PRIMARY KEY,
                    data JSONB,
                    updated_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS paid_marks (
                    user_id BIGINT NOT NULL,
                    person_key TEXT NOT NULL,
                    paid_at TIMESTAMP DEFAULT NOW(),
                    PRIMARY KEY (user_id, person_key)
                );
            """)
        conn.commit()

def db_get_vedomosti(uid):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT vedomist_type, dilinitsa, period, source_file, rows FROM vedomosti WHERE user_id=%s ORDER BY created_at", (uid,))
            rows = cur.fetchall()
    return [{"vedomist_type":r[0],"dilinitsa":r[1],"period":r[2],"source_file":r[3],"rows":r[4]} for r in rows]

def db_save_vedomist(uid, v):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO vedomosti (user_id, vedomist_type, dilinitsa, period, source_file, rows)
                VALUES (%s,%s,%s,%s,%s,%s)
                ON CONFLICT (user_id, vedomist_type, dilinitsa, period)
                DO UPDATE SET rows=EXCLUDED.rows, source_file=EXCLUDED.source_file, created_at=NOW()
            """, (uid, v.get("vedomist_type"), v.get("dilinitsa"), v.get("period"), v.get("source_file"), Json(v.get("rows",[]))))
        conn.commit()

def db_clear_vedomosti(uid):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM vedomosti WHERE user_id=%s", (uid,))
        conn.commit()

def db_get_phones(uid):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT data FROM phone_books WHERE user_id=%s", (uid,))
            row = cur.fetchone()
    return row[0] if row else {}

def db_save_phones(uid, data):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO phone_books (user_id, data, updated_at)
                VALUES (%s,%s,NOW())
                ON CONFLICT (user_id) DO UPDATE SET data=EXCLUDED.data, updated_at=NOW()
            """, (uid, Json(data)))
        conn.commit()

# ── Paid marks DB ────────────────────────────────────────────────────────────
def db_get_paid(uid):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT person_key FROM paid_marks WHERE user_id=%s", (uid,))
            return {row[0] for row in cur.fetchall()}

def db_mark_paid(uid, key):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO paid_marks (user_id, person_key) VALUES (%s,%s) ON CONFLICT DO NOTHING", (uid, key))
        conn.commit()

def db_unmark_paid(uid, key):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM paid_marks WHERE user_id=%s AND person_key=%s", (uid, key))
        conn.commit()

def db_clear_paid(uid):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM paid_marks WHERE user_id=%s", (uid,))
        conn.commit()

def make_person_key(r):
    # Use account number as primary key, fallback to name+addr
    acc = str(r.get("account","") or r.get("passport","") or "").strip()
    if acc:
        return "ACC|" + acc
    return r["name"] + "|" + norm_addr(r.get("address",""))

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm_addr(addr: str) -> str:
    return re.sub(r'[^А-ЯІЇЄҐа-яіїєґA-Za-z0-9]', '',
        addr.upper()
        .replace('ВУЛ.','').replace('ВУЛИЦЯ','').replace('СОШИЧНЕ','')
        .replace(';;','').replace('Б/Н','').replace('М.КАМІНЬ-КАШИРСЬКИЙ','')
        .replace('С.ЩИТИНЬ','').replace('С.СОШИЧНЕ',''))

def find_phone(phone_book, addr):
    if not phone_book: return ''
    key = norm_addr(addr)
    if not key: return ''
    if key in phone_book: return phone_book[key]
    for k, v in phone_book.items():
        if len(key) > 5 and (k.startswith(key) or key.startswith(k)):
            return v
    return ''

def parse_date(s):
    m = re.match(r'(\d{1,2})[.\-/](\d{1,2})', str(s))
    if not m: return None
    try: return datetime(datetime.now().year, int(m.group(2)), int(m.group(1)))
    except: return None

def is_payable_by(pay_date, limit_date):
    d = parse_date(pay_date)
    if not d: return False
    return d <= limit_date

def fmt_hrn(n):
    try: return f"{float(n):,.2f} ₴".replace(',', ' ')
    except: return str(n)

def esc(text):
    BACKSLASH = chr(92)
    chars = "_[]()~`>#+-=|{}.!"
    result = []
    for c in str(text):
        if c in chars:
            result.append(BACKSLASH + c)
        else:
            result.append(c)
    return "".join(result)

def parse_xlsx(data: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    phone_map = {}
    for row in ws.iter_rows(values_only=True):
        col1 = str(row[1] if len(row) > 1 else '').strip()
        col3 = str(row[3] if len(row) > 3 else '').strip()
        if 'вул.' in col1.lower() and col3 and col3 != 'None':
            key = norm_addr(col1)
            if key and key not in phone_map:
                phone_map[key] = col3
    return phone_map

def extract_json(text):
    text = re.sub(r'^```(?:json)?\s*', '', text.strip())
    text = re.sub(r'\s*```$', '', text.strip()).strip()
    try: return json.loads(text)
    except: pass
    s, e = text.find('{'), text.rfind('}')
    if s != -1 and e > s:
        try: return json.loads(text[s:e+1])
        except: pass
    return None

def build_all_rows(vedomosti, phone_book):
    grouped = {}
    for v in vedomosti:
        for r in v.get("rows", []):
            key = make_person_key(r)
            if key not in grouped:
                grouped[key] = {
                    "name": r["name"],
                    "address": r.get("address",""),
                    "passport": r.get("passport",""),
                    "account": r.get("account",""),
                    "phone": find_phone(phone_book, r.get("address","")),
                    "veds": [],
                    "_key": key
                }
            grouped[key]["veds"].append({
                "type": v.get("vedomist_type","?"),
                "dil":  v.get("dilinitsa","?"),
                "sum":  r.get("sum", 0),
                "pay_date": r.get("pay_date","?")
            })
    return list(grouped.values())

def format_rows(rows, phone_book, paid=None):
    if paid is None: paid = set()
    chunks = []
    current = ""
    for i, r in enumerate(rows, 1):
        key   = make_person_key(r)
        is_paid = key in paid
        addr  = r["address"].replace("Сошичне, ","").replace(";;","").strip()
        phone = r.get("phone","") or find_phone(phone_book, r["address"])
        veds_str = " ".join([f"В{v['type']}·{v['dil']}" for v in r["veds"]])
        pays_str = " / ".join([f"{fmt_hrn(v.get('sum',0))} · {v['pay_date']}" for v in r["veds"]])

        account = r.get("account","")
        acc_str = f"   🔢 {account}\n" if account else ""
        if is_paid:
            line = str(esc(i)) + "\\." + " ✅ *" + esc(r["name"]) + "*" + chr(10)
            line += "   📍 " + esc(addr) + chr(10)
            line += "   ✔️ *Виплачено*" + chr(10)
            line += "   💰 " + esc(pays_str) + chr(10)
            if account:
                line += "   🔢 " + esc(account) + chr(10)
            line += chr(10)
        else:
            line = str(esc(i)) + "\\." + " " + esc(r["name"]) + chr(10)
            line += "   📍 " + esc(addr) + chr(10)
            if phone:
                line += "   📞 " + esc(phone) + chr(10)
            line += "   📋 " + esc(veds_str) + chr(10)
            line += "   💰 " + esc(pays_str) + chr(10)
            line += "   🆔 " + esc(r["passport"]) + chr(10)
            if account:
                line += "   🔢 " + esc(account) + chr(10)
            line += chr(10)