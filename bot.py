import os, json, base64, re, io
from datetime import datetime
import httpx
import openpyxl
import psycopg2
from psycopg2.extras import Json
from telegram import Update
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

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm_addr(addr):
    return re.sub(r'[^А-ЯІЇЄҐа-яіїєґA-Za-z0-9]', '',
        addr.upper()
        .replace('ВУЛ.','').replace('ВУЛИЦЯ','').replace('СОШИЧНЕ','')
        .replace(';;','').replace('Б/Н','').replace('М.КАМІНЬ-КАШИРСЬКИЙ','')
        .replace('С.ЩИТИНЬ','').replace('С.СОШИЧНЕ',''))

def make_person_key(r):
    acc = str(r.get("account","") or "").strip()
    if acc:
        return "ACC|" + acc
    return r["name"] + "|" + norm_addr(r.get("address",""))

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
    try: return "{:,.2f} грн".format(float(n)).replace(',', ' ')
    except: return str(n)

def parse_xlsx(data):
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
                    "veds": []
                }
            grouped[key]["veds"].append({
                "type": v.get("vedomist_type","?"),
                "dil":  v.get("dilinitsa","?"),
                "sum":  r.get("sum", 0),
                "pay_date": r.get("pay_date","?")
            })
    return list(grouped.values())

def format_rows(rows, phone_book, paid=None):
    if paid is None:
        paid = set()
    chunks = []
    current = ""
    for i, r in enumerate(rows, 1):
        key     = make_person_key(r)
        is_paid = key in paid
        addr    = r["address"].replace("Сошичне, ","").replace(";;","").strip()
        phone   = r.get("phone","") or find_phone(phone_book, r["address"])
        account = r.get("account","")
        veds_str = " ".join(["В" + v["type"] + "·" + v["dil"] for v in r["veds"]])
        pays_str = " / ".join([fmt_hrn(v.get("sum",0)) + " · " + v["pay_date"] for v in r["veds"]])

        if is_paid:
            line = str(i) + ". ✅ " + r["name"] + "\n"
            line += "   📍 " + addr + "\n"
            line += "   ✔️ Виплачено\n"
            line += "   💰 " + pays_str + "\n"
            if account:
                line += "   🔢 " + account + "\n"
        else:
            line = str(i) + ". " + r["name"] + "\n"
            line += "   📍 " + addr + "\n"
            if phone:
                line += "   📞 " + phone + "\n"
            line += "   📋 " + veds_str + "\n"
            line += "   💰 " + pays_str + "\n"
            line += "   🪪 " + r["passport"] + "\n"
            if account:
                line += "   🔢 " + account + "\n"
        line += "\n"

        if len(current) + len(line) > 3800:
            chunks.append(current)
            current = line
        else:
            current += line
    if current:
        chunks.append(current)
    return chunks

def find_by_account_or_name(all_rows, query):
    q = query.strip()
    by_acc = [r for r in all_rows if r.get("account","") == q]
    if by_acc: return by_acc
    return [r for r in all_rows if q.upper() in r["name"].upper()]

# ── Claude API ────────────────────────────────────────────────────────────────
async def ask_claude(b64):
    prompt = (
        "Це відомість ПФУ на виплату пенсій. Зчитай всі рядки таблиці з УСІХ сторінок.\n"
        "Поверни ТІЛЬКИ валідний JSON без markdown, без пояснень:\n"
        '{"vedomist_type":"99","period":"Червень 2026 01","dilinitsa":"50","rows":[{"account":"001256811781","name":"БАЛЕЦЬКИЙ ВОЛОДИМИР ПОРФИРІЙОВИЧ","address":"Сошичне, МИРУ, 94;;","sum":2301.28,"passport":"AC 000393421","pay_date":"04.06"}]}\n'
        "- account: номер особового рахунку з першої колонки\n"
        "- vedomist_type: число з рядка Тип виплати\n"
        "- dilinitsa: перший номер після назви села\n"
        "- period: текст після слова період\n"
        "- name: ПІБ ВЕЛИКИМИ ЛІТЕРАМИ\n"
        "- address: адреса як є\n"
        "- sum: число\n"
        "- pay_date: DD.MM"
    )
    async with httpx.AsyncClient(timeout=180) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01",
                "anthropic-beta": "pdfs-2024-09-25",
                "Content-Type": "application/json",
            },
            json={
                "model": "claude-haiku-4-5-20251001",
                "max_tokens": 8192,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            }
        )
        resp.raise_for_status()
        data = resp.json()
        text = "".join(c.get("text","") for c in data.get("content",[]))
        parsed = extract_json(text)
        if not parsed:
            raise ValueError("JSON не знайдено. Відповідь: " + text[:300])
        if not isinstance(parsed.get("rows"), list):
            parsed["rows"] = []
        return parsed

# ── Handlers ──────────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привіт! Бот для обробки відомостей ПФУ.\n\n"
        "📄 Надішли PDF — додається до загального списку\n"
        "📞 Надішли XLSX — завантажує телефонну книгу\n\n"
        "Команди:\n"
        "/status — скільки відомостей і людей\n"
        "/list — зведений список всіх\n"
        "/today 07.06 — виплати до вказаної дати\n"
        "/search Шв — пошук (2+ літери)\n"
        "/multi — отримувачі з 2+ виплатами\n"
        "/clear — очистити відомості\n"
        "/phones — статус телефонної книги\n\n"
        "Відмітки виплат:\n"
        "/paid 001256811781 — відмітити виплачено\n"
        "/unpaid 001256811781 — зняти позначку\n"
        "/unpaid_list — список невиплачених\n"
        "/clear_paid — зняти всі позначки"
    )

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    rows = build_all_rows(veds, phones)
    paid = db_get_paid(uid)
    total_sum = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in rows)
    unpaid_count = sum(1 for r in rows if make_person_key(r) not in paid)
    ved_list = "\n".join(["  • В" + v.get("vedomist_type","?") + "·Діл." + v.get("dilinitsa","?") + " — " + str(len(v.get("rows",[]))) + " ос. (" + v.get("period","") + ")" for v in veds])
    await update.message.reply_text(
        "📊 Статус:\n" + ved_list + "\n\n"
        "👥 Унікальних осіб: " + str(len(rows)) + "\n"
        "✅ Виплачено: " + str(len(rows) - unpaid_count) + "\n"
        "⏳ Залишилось: " + str(unpaid_count) + "\n"
        "💰 Загальна сума: " + fmt_hrn(total_sum) + "\n"
        "📞 Телефонна книга: " + str(len(phones)) + " адрес"
    )

async def cmd_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    rows = build_all_rows(veds, phones)
    rows.sort(key=lambda r: r["veds"][0].get("pay_date","99.99"))
    paid = db_get_paid(uid)
    unpaid_count = sum(1 for r in rows if make_person_key(r) not in paid)
    await update.message.reply_text("📋 Зведений список · " + str(len(rows)) + " осіб · не виплачено: " + str(unpaid_count) + "\n" + "─"*28)
    for chunk in format_rows(rows, phones, paid):
        await update.message.reply_text(chunk)

async def cmd_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    if not ctx.args:
        await update.message.reply_text("⚠️ Вкажи дату: /today 07.06")
        return
    limit = parse_date(ctx.args[0])
    if not limit:
        await update.message.reply_text("⚠️ Невірний формат. Приклад: /today 07.06")
        return
    all_rows = build_all_rows(veds, phones)
    filtered = []
    for r in all_rows:
        payable = [v for v in r["veds"] if is_payable_by(v["pay_date"], limit)]
        if payable:
            filtered.append(dict(r, veds=payable))
    if not filtered:
        await update.message.reply_text("📭 Немає виплат до " + ctx.args[0])
        return
    filtered.sort(key=lambda r: r["veds"][0].get("pay_date","99.99"))
    paid = db_get_paid(uid)
    total = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in filtered)
    unpaid = sum(1 for r in filtered if make_person_key(r) not in paid)
    await update.message.reply_text("📅 Виплати до " + ctx.args[0] + " · " + str(len(filtered)) + " осіб · не виплачено: " + str(unpaid) + "\n💰 " + fmt_hrn(total) + "\n" + "─"*28)
    for chunk in format_rows(filtered, phones, paid):
        await update.message.reply_text(chunk)

async def cmd_search(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    if not ctx.args:
        await update.message.reply_text("⚠️ Вкажи частину прізвища: /search Шв")
        return
    query = " ".join(ctx.args).upper().strip()
    if len(query) < 2:
        await update.message.reply_text("⚠️ Мінімум 2 літери")
        return
    found = [r for r in build_all_rows(veds, phones) if query in r["name"].upper()]
    if not found:
        await update.message.reply_text("🔍 Нічого не знайдено: " + query)
        return
    paid = db_get_paid(uid)
    await update.message.reply_text("🔍 '" + query + "' · " + str(len(found)) + " осіб\n" + "─"*28)
    for chunk in format_rows(found, phones, paid):
        await update.message.reply_text(chunk)

async def cmd_multi(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    all_rows = build_all_rows(veds, phones)
    multi = [r for r in all_rows if len(r["veds"]) >= 2]
    if not multi:
        await update.message.reply_text("📭 Немає отримувачів з двома і більше виплатами.")
        return
    multi.sort(key=lambda r: r["name"])
    paid = db_get_paid(uid)
    total = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in multi)
    await update.message.reply_text("👥 Отримувачів з 2+ виплатами: " + str(len(multi)) + "\n💰 Загальна сума: " + fmt_hrn(total) + "\n" + "─"*28)
    for chunk in format_rows(multi, phones, paid):
        await update.message.reply_text(chunk)

async def cmd_paid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає.")
        return
    if not ctx.args:
        await update.message.reply_text("⚠️ Вкажи номер рахунку або прізвище:\n/paid 001256811781\n/paid Шворак")
        return
    query = " ".join(ctx.args).strip()
    all_rows = build_all_rows(veds, phones)
    found = find_by_account_or_name(all_rows, query)
    if not found:
        await update.message.reply_text("🔍 Не знайдено: " + query)
        return
    if len(found) > 1:
        names = "\n".join([str(i+1) + ". " + r.get("account","") + " " + r["name"] for i, r in enumerate(found)])
        await update.message.reply_text("Знайдено кілька — вкажи номер рахунку:\n" + names)
        return
    r = found[0]
    db_mark_paid(uid, make_person_key(r))
    await update.message.reply_text("✅ Виплачено: " + r["name"] + "\n   🔢 " + r.get("account",""))

async def cmd_unpaid_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає.")
        return
    if not ctx.args:
        await update.message.reply_text("⚠️ Вкажи номер рахунку або прізвище:\n/unpaid 001256811781")
        return
    query = " ".join(ctx.args).strip()
    all_rows = build_all_rows(veds, phones)
    found = find_by_account_or_name(all_rows, query)
    if not found:
        await update.message.reply_text("🔍 Не знайдено: " + query)
        return
    if len(found) > 1:
        names = "\n".join([str(i+1) + ". " + r.get("account","") + " " + r["name"] for i, r in enumerate(found)])
        await update.message.reply_text("Знайдено кілька — вкажи номер рахунку:\n" + names)
        return
    r = found[0]
    db_unmark_paid(uid, make_person_key(r))
    await update.message.reply_text("↩️ Знято позначку: " + r["name"] + "\n   🔢 " + r.get("account",""))

async def cmd_unpaid_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    phones = db_get_phones(uid)
    if not veds:
        await update.message.reply_text("📭 Відомостей немає.")
        return
    all_rows = build_all_rows(veds, phones)
    paid = db_get_paid(uid)
    unpaid = [r for r in all_rows if make_person_key(r) not in paid]
    unpaid.sort(key=lambda r: r["veds"][0].get("pay_date","99.99"))
    if not unpaid:
        await update.message.reply_text("🎉 Всі виплати здійснено!")
        return
    total = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in unpaid)
    await update.message.reply_text("⏳ Не виплачено · " + str(len(unpaid)) + " осіб · " + fmt_hrn(total) + "\n" + "─"*28)
    for chunk in format_rows(unpaid, phones, paid):
        await update.message.reply_text(chunk)

async def cmd_clear_paid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    db_clear_paid(uid)
    await update.message.reply_text("↩️ Всі позначки виплат знято")

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    veds = db_get_vedomosti(uid)
    db_clear_vedomosti(uid)
    await update.message.reply_text("🗑 Очищено " + str(len(veds)) + " відомостей. Телефонна книга збережена.")

async def cmd_phones(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    phones = db_get_phones(uid)
    if phones:
        await update.message.reply_text("📞 Телефонна книга: " + str(len(phones)) + " адрес")
    else:
        await update.message.reply_text("📞 Телефонна книга не завантажена\nНадішли XLSX файл.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    uid = update.effective_user.id
    is_xlsx = (doc.mime_type in (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel"
    ) or doc.file_name.lower().endswith(('.xlsx','.xls')))

    if is_xlsx:
        msg = await update.message.reply_text("⏳ Завантажую телефонну книгу...")
        try:
            file = await ctx.bot.get_file(doc.file_id)
            data = await file.download_as_bytearray()
            phone_map = parse_xlsx(bytes(data))
            db_save_phones(uid, phone_map)
            await msg.edit_text("✅ Телефонна книга: " + str(len(phone_map)) + " адрес збережено")
        except Exception as e:
            await msg.edit_text("❌ Помилка XLSX: " + str(e)[:300])
        return

    if doc.mime_type != "application/pdf":
        await update.message.reply_text("⚠️ Надішли PDF відомість або XLSX телефонну книгу")
        return

    msg = await update.message.reply_text("⏳ Завантажую " + doc.file_name + "...")
    try:
        file = await ctx.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        b64 = base64.b64encode(pdf_bytes).decode()
        await msg.edit_text("🤖 Обробляю через Claude (" + str(len(pdf_bytes)//1024) + " KB)...")
        result = await ask_claude(b64)
        result["source_file"] = doc.file_name
        rows = result.get("rows", [])
        if not rows:
            await msg.edit_text("⚠️ Рядків не знайдено")
            return
        db_save_vedomist(uid, result)
        veds = db_get_vedomosti(uid)
        phones = db_get_phones(uid)
        total_people = len(build_all_rows(veds, phones))
        ved = result.get("vedomist_type","?")
        dil = result.get("dilinitsa","?")
        per = result.get("period","?")
        await msg.edit_text(
            "✅ Збережено: В" + ved + "·Діл." + dil + " · " + per + " · " + str(len(rows)) + " осіб\n\n"
            "📊 Всього відомостей: " + str(len(veds)) + "\n"
            "👥 Унікальних осіб: " + str(total_people) + "\n\n"
            "Команди: /list · /today 07.06 · /status"
        )
    except Exception as e:
        await msg.edit_text("❌ Помилка: " + str(e)[:500])

async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    text = update.message.text.strip().upper()
    if len(text) >= 2:
        veds = db_get_vedomosti(uid)
        phones = db_get_phones(uid)
        if veds:
            found = [r for r in build_all_rows(veds, phones) if text in r["name"].upper()]
            if found:
                paid = db_get_paid(uid)
                await update.message.reply_text("🔍 '" + update.message.text.strip() + "' · " + str(len(found)) + " осіб\n" + "─"*28)
                for chunk in format_rows(found, phones, paid):
                    await update.message.reply_text(chunk)
                return
    await update.message.reply_text(
        "📄 Надішли PDF або XLSX\n\n"
        "Команди: /list · /today 07.06 · /status · /clear\n"
        "/search Шв або просто напиши прізвище"
    )

def main():
    init_db()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",       start))
    app.add_handler(CommandHandler("status",      cmd_status))
    app.add_handler(CommandHandler("list",        cmd_list))
    app.add_handler(CommandHandler("today",       cmd_today))
    app.add_handler(CommandHandler("search",      cmd_search))
    app.add_handler(CommandHandler("multi",       cmd_multi))
    app.add_handler(CommandHandler("clear",       cmd_clear))
    app.add_handler(CommandHandler("phones",      cmd_phones))
    app.add_handler(CommandHandler("paid",        cmd_paid))
    app.add_handler(CommandHandler("unpaid",      cmd_unpaid_cmd))
    app.add_handler(CommandHandler("unpaid_list", cmd_unpaid_list))
    app.add_handler(CommandHandler("clear_paid",  cmd_clear_paid))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot started with DB")
    app.run_polling(close_loop=False)

if __name__ == "__main__":
    main()
