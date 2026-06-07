import os, json, base64, re, io
from datetime import datetime
import httpx
import openpyxl
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]

# Storage per user
# user_data[uid] = {
#   "vedomosti": [...],   # list of parsed vedomist dicts
#   "phone_book": {}      # norm_addr -> phone
# }
user_data = {}

def get_ud(uid):
    if uid not in user_data:
        user_data[uid] = {"vedomosti": [], "phone_book": {}}
    return user_data[uid]

# ── Address normalization ─────────────────────────────────────────────────────
def norm_addr(addr: str) -> str:
    return re.sub(r'[^А-ЯІЇЄҐа-яіїєґA-Za-z0-9]', '',
        addr.upper()
        .replace('ВУЛ.','').replace('ВУЛИЦЯ','').replace('СОШИЧНЕ','')
        .replace(';;','').replace('Б/Н','').replace('М.КАМІНЬ-КАШИРСЬКИЙ','')
        .replace('С.ЩИТИНЬ','').replace('С.СОШИЧНЕ',''))

def find_phone(phone_book: dict, addr: str) -> str:
    if not phone_book: return ''
    key = norm_addr(addr)
    if not key: return ''
    if key in phone_book: return phone_book[key]
    for k, v in phone_book.items():
        if len(key) > 5 and (k.startswith(key) or key.startswith(k)):
            return v
    return ''

# ── Date helpers ──────────────────────────────────────────────────────────────
def parse_date(s: str):
    """Parse DD.MM -> date object with current year"""
    m = re.match(r'(\d{1,2})[.\-/](\d{1,2})', str(s))
    if not m: return None
    try:
        return datetime(datetime.now().year, int(m.group(2)), int(m.group(1)))
    except: return None

def is_payable_by(pay_date: str, limit_date: datetime) -> bool:
    d = parse_date(pay_date)
    if not d: return False
    return d <= limit_date

# ── XLSX parser ───────────────────────────────────────────────────────────────
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

# ── JSON extractor ────────────────────────────────────────────────────────────
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

def fmt_hrn(n):
    try: return f"{float(n):,.2f} ₴".replace(',', ' ')
    except: return str(n)

# ── Claude API ────────────────────────────────────────────────────────────────
async def ask_claude(b64: str) -> dict:
    prompt = """Це відомість ПФУ на виплату пенсій. Зчитай всі рядки таблиці з УСІХ сторінок.
Поверни ТІЛЬКИ валідний JSON без markdown, без пояснень:
{"vedomist_type":"99","period":"Червень 2026 01","dilinitsa":"50","rows":[{"name":"БАЛЕЦЬКИЙ ВОЛОДИМИР ПОРФИРІЙОВИЧ","address":"Сошичне, МИРУ, 94;;","sum":2301.28,"passport":"AC 000393421","pay_date":"04.06"}]}"""

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
            raise ValueError(f"JSON не знайдено. Відповідь: {text[:300]}")
        if not isinstance(parsed.get("rows"), list):
            parsed["rows"] = []
        return parsed

# ── Build grouped rows from all vedomosti ────────────────────────────────────
def build_all_rows(vedomosti: list, phone_book: dict) -> list:
    """Merge all vedomosti, group by name+address, collect all veds per person"""
    grouped = {}
    for v in vedomosti:
        for r in v.get("rows", []):
            key = r["name"] + "|" + norm_addr(r.get("address",""))
            if key not in grouped:
                grouped[key] = {
                    "name": r["name"],
                    "address": r.get("address",""),
                    "passport": r.get("passport",""),
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

def format_rows(rows: list, phone_book: dict) -> list:
    """Format rows into Telegram message chunks"""
    chunks = []
    current = ""
    for i, r in enumerate(rows, 1):
        addr = r["address"].replace("Сошичне, ","").replace(";;","").strip()
        phone = r.get("phone","") or find_phone(phone_book, r["address"])
        veds_str = " ".join([f"В{v['type']}·{v['dil']}" for v in r["veds"]])
        dates_str = " / ".join([f"{v['pay_date']}" for v in r["veds"]])
        total_sum = sum(float(v.get("sum",0)) for v in r["veds"])

        line = f"{i}. {r['name']}\n   📍 {addr}\n"
        if phone:
            line += f"   📞 {phone}\n"
        line += f"   📋 {veds_str}\n"
        line += f"   💰 {fmt_hrn(total_sum)} · 📅 {dates_str}\n   🪪 {r['passport']}\n\n"

        if len(current) + len(line) > 3800:
            chunks.append(current)
            current = line
        else:
            current += line

    if current:
        chunks.append(current)
    return chunks

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
        "/clear — очистити всі відомості\n"
        "/phones — статус телефонної книги"
    )

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = get_ud(uid)
    veds = ud["vedomosti"]
    if not veds:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    rows = build_all_rows(veds, ud["phone_book"])
    total_sum = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in rows)
    ved_list = "\n".join([f"  • В{v.get('vedomist_type','?')}·Діл.{v.get('dilinitsa','?')} — {len(v.get('rows',[]))} ос. ({v.get('period','')})" for v in veds])
    await update.message.reply_text(
        f"📊 Статус:\n{ved_list}\n\n"
        f"👥 Всього унікальних осіб: {len(rows)}\n"
        f"💰 Загальна сума: {fmt_hrn(total_sum)}\n"
        f"📞 Телефонна книга: {len(ud['phone_book'])} адрес"
    )

async def cmd_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = get_ud(uid)
    if not ud["vedomosti"]:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return
    rows = build_all_rows(ud["vedomosti"], ud["phone_book"])
    rows.sort(key=lambda r: r["veds"][0].get("pay_date","99.99"))
    header = f"📋 Зведений список · {len(rows)} осіб\n{'─'*28}\n"
    chunks = format_rows(rows, ud["phone_book"])
    await update.message.reply_text(header)
    for chunk in chunks:
        await update.message.reply_text(chunk)

async def cmd_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = get_ud(uid)
    if not ud["vedomosti"]:
        await update.message.reply_text("📭 Відомостей немає. Надішли PDF.")
        return

    # Parse date from command args
    args = ctx.args
    if not args:
        await update.message.reply_text("⚠️ Вкажи дату: /today 07.06")
        return
    limit = parse_date(args[0])
    if not limit:
        await update.message.reply_text("⚠️ Невірний формат дати. Приклад: /today 07.06")
        return

    all_rows = build_all_rows(ud["vedomosti"], ud["phone_book"])
    # Filter rows that have at least one ved payable by limit date
    filtered = []
    for r in all_rows:
        payable_veds = [v for v in r["veds"] if is_payable_by(v["pay_date"], limit)]
        if payable_veds:
            filtered.append({**r, "veds": payable_veds})

    if not filtered:
        await update.message.reply_text(f"📭 Немає виплат до {args[0]}")
        return

    filtered.sort(key=lambda r: r["veds"][0].get("pay_date","99.99"))
    total = sum(sum(float(v.get("sum",0)) for v in r["veds"]) for r in filtered)
    header = f"📅 Виплати до {args[0]} · {len(filtered)} осіб · {fmt_hrn(total)}\n{'─'*28}\n"
    chunks = format_rows(filtered, ud["phone_book"])
    await update.message.reply_text(header)
    for chunk in chunks:
        await update.message.reply_text(chunk)

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = get_ud(uid)
    count = len(ud["vedomosti"])
    ud["vedomosti"] = []
    await update.message.reply_text(f"🗑 Очищено {count} відомостей. Телефонна книга збережена.")

async def cmd_phones(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = get_ud(uid)
    n = len(ud["phone_book"])
    if n:
        await update.message.reply_text(f"📞 Телефонна книга: {n} адрес завантажено")
    else:
        await update.message.reply_text("📞 Телефонна книга не завантажена\nНадішли XLSX файл.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    uid = update.effective_user.id
    ud = get_ud(uid)

    # XLSX
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
            ud["phone_book"] = phone_map
            await msg.edit_text(f"✅ Телефонна книга: {len(phone_map)} адрес завантажено")
        except Exception as e:
            await msg.edit_text(f"❌ Помилка XLSX: {str(e)[:300]}")
        return

    # PDF
    if doc.mime_type != "application/pdf":
        await update.message.reply_text("⚠️ Надішли PDF відомість або XLSX телефонну книгу")
        return

    msg = await update.message.reply_text(f"⏳ Завантажую {doc.file_name}...")
    try:
        file = await ctx.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        b64 = base64.b64encode(pdf_bytes).decode()

        await msg.edit_text(f"🤖 Обробляю через Claude ({len(pdf_bytes)//1024} KB)...")

        result = await ask_claude(b64)
        result["source_file"] = doc.file_name
        rows = result.get("rows", [])

        if not rows:
            await msg.edit_text("⚠️ Рядків не знайдено")
            return

        # Check if this vedomist already loaded (same type+dilinitsa+period)
        key = f"{result.get('vedomist_type')}|{result.get('dilinitsa')}|{result.get('period')}"
        existing = next((i for i, v in enumerate(ud["vedomosti"])
                        if f"{v.get('vedomist_type')}|{v.get('dilinitsa')}|{v.get('period')}" == key), None)

        if existing is not None:
            ud["vedomosti"][existing] = result
            status = "♻️ Оновлено"
        else:
            ud["vedomosti"].append(result)
            status = "✅ Додано"

        ved = result.get("vedomist_type","?")
        dil = result.get("dilinitsa","?")
        per = result.get("period","?")
        total_veds = len(ud["vedomosti"])
        total_people = len(build_all_rows(ud["vedomosti"], ud["phone_book"]))

        await msg.edit_text(
            f"{status}: В{ved}·Діл.{dil} · {per} · {len(rows)} осіб\n\n"
            f"📊 Всього відомостей: {total_veds}\n"
            f"👥 Унікальних осіб: {total_people}\n\n"
            f"Команди: /list · /today 07.06 · /status"
        )

    except Exception as e:
        await msg.edit_text(f"❌ Помилка: {str(e)[:500]}")

async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📄 Надішли PDF відомість або XLSX телефонну книгу\n\n"
        "Команди: /list · /today 07.06 · /status · /clear"
    )

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("list",   cmd_list))
    app.add_handler(CommandHandler("today",  cmd_today))
    app.add_handler(CommandHandler("clear",  cmd_clear))
    app.add_handler(CommandHandler("phones", cmd_phones))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot started")
    app.run_polling(close_loop=False)

if __name__ == "__main__":
    main()
