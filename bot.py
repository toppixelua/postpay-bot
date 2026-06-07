import os, json, base64, re
import httpx
import openpyxl
import io
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]

# In-memory phone book storage per user
phone_books = {}  # user_id -> {norm_addr -> phone}

def norm_addr(addr: str) -> str:
    return re.sub(r'[^А-ЯІЇЄҐа-яіїєґA-Za-z0-9]', '', addr.upper()
        .replace('ВУЛ.','').replace('ВУЛИЦЯ','').replace('СОШИЧНЕ','')
        .replace(';;','').replace('Б/Н',''))

def find_phone(phone_book: dict, addr: str) -> str:
    if not phone_book:
        return ''
    key = norm_addr(addr)
    if not key:
        return ''
    if key in phone_book:
        return phone_book[key]
    for k, v in phone_book.items():
        if len(key) > 5 and (k.startswith(key) or key.startswith(k)):
            return v
    return ''

def parse_xlsx(data: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    phone_map = {}
    for row in ws.iter_rows(values_only=True):
        col1 = str(row[1] or '').strip()
        col3 = str(row[3] or '').strip()
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

def fmt_hrn(n):
    try: return f"{float(n):,.2f} ₴".replace(',', ' ')
    except: return str(n)

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

# ── Handlers ──────────────────────────────────────────────────────────────────

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привіт! Я бот для обробки відомостей ПФУ.\n\n"
        "📄 Надішли PDF відомість — отримаєш список виплат\n"
        "📞 Надішли XLSX телефонну книгу — бот підтягуватиме телефони\n\n"
        "Команди:\n"
        "/start — це повідомлення\n"
        "/phones — статус телефонної книги"
    )

async def phones_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    book = phone_books.get(uid, {})
    if book:
        await update.message.reply_text(f"📞 Телефонна книга: {len(book)} адрес завантажено")
    else:
        await update.message.reply_text("📞 Телефонна книга не завантажена\n\nНадішли XLSX файл з адресами і телефонами.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    uid = update.effective_user.id

    # XLSX — телефонна книга
    if doc.mime_type in ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          "application/vnd.ms-excel") or doc.file_name.endswith(('.xlsx','.xls')):
        msg = await update.message.reply_text("⏳ Завантажую телефонну книгу...")
        try:
            file = await ctx.bot.get_file(doc.file_id)
            data = await file.download_as_bytearray()
            phone_map = parse_xlsx(bytes(data))
            phone_books[uid] = phone_map
            await msg.edit_text(f"✅ Телефонна книга завантажена: {len(phone_map)} адрес")
        except Exception as e:
            await msg.edit_text(f"❌ Помилка XLSX: {str(e)[:300]}")
        return

    # PDF — відомість
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
        rows   = result.get("rows", [])
        ved    = result.get("vedomist_type", "?")
        dil    = result.get("dilinitsa", "?")
        per    = result.get("period", "?")

        if not rows:
            await msg.edit_text("⚠️ Рядків не знайдено у відомості")
            return

        book = phone_books.get(uid, {})
        has_phones = bool(book)

        header = f"📋 Відомість тип {ved} · Діл. {dil}\n📅 {per}\n👥 {len(rows)} осіб\n{'─'*28}\n"
        chunks = []
        current = header

        for i, r in enumerate(rows, 1):
            name     = r.get("name","?")
            addr     = r.get("address","?").replace("Сошичне, ","").replace(";;","").strip()
            summa    = fmt_hrn(r.get("sum", 0))
            date     = r.get("pay_date","?")
            passport = r.get("passport","")
            phone    = find_phone(book, r.get("address","")) if has_phones else ""

            line = f"{i}. {name}\n   📍 {addr}\n"
            if phone:
                line += f"   📞 {phone}\n"
            line += f"   💰 {summa} · 📅 {date}\n   🪪 {passport}\n\n"

            if len(current) + len(line) > 3800:
                chunks.append(current)
                current = line
            else:
                current += line

        if current:
            chunks.append(current)

        await msg.delete()
        for chunk in chunks:
            await update.message.reply_text(chunk)

        total = sum(float(r.get("sum", 0)) for r in rows)
        phones_found = sum(1 for r in rows if find_phone(book, r.get("address",""))) if has_phones else 0
        summary = f"✅ Готово!\n💰 Загальна сума: {fmt_hrn(total)}\n📄 {doc.file_name}"
        if has_phones:
            summary += f"\n📞 Телефонів знайдено: {phones_found}/{len(rows)}"
        await update.message.reply_text(summary)

    except Exception as e:
        await msg.edit_text(f"❌ Помилка: {str(e)[:500]}")

async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📄 Надішли PDF відомість або XLSX телефонну книгу")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("phones", phones_status))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot started")
    app.run_polling(close_loop=False)

if __name__ == "__main__":
    main()
