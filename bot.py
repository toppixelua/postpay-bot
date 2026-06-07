import os, json, base64, re, asyncio
import httpx
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]

def extract_json(text):
    m = re.search(r'```(?:json)?\s*([\s\S]*?)```', text)
    if m:
        try: return json.loads(m.group(1).strip())
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
Поверни ТІЛЬКИ JSON без тексту навколо:
{"vedomist_type":"99","period":"Червень 2026 01","dilinitsa":"50","rows":[{"name":"БАЛЕЦЬКИЙ ВОЛОДИМИР ПОРФИРІЙОВИЧ","address":"Сошичне, МИРУ, 94;;","sum":2301.28,"passport":"AC 000393421","pay_date":"04.06"}]}
- vedomist_type: число з рядка "Тип виплати:"
- dilinitsa: перший номер після назви села в рядку "Найменування виплатного обєкта..."
- period: текст після слова "період"
- name: ПІБ ВЕЛИКИМИ ЛІТЕРАМИ
- address: адреса як є
- sum: число
- pay_date: дата DD.MM"""

    async with httpx.AsyncClient(timeout=120) as client:
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
                "max_tokens": 4096,
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
            raise ValueError(f"Не знайдено JSON. Claude: {text[:200]}")
        if not isinstance(parsed.get("rows"), list):
            parsed["rows"] = []
        return parsed

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привіт! Надішли PDF відомість ПФУ — я розпізнаю всі рядки і виведу список виплат.\n\n"
        "Можна надсилати кілька файлів підряд."
    )

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        await update.message.reply_text("⚠️ Надішли PDF файл")
        return

    msg = await update.message.reply_text(f"⏳ Обробляю {doc.file_name}...")

    try:
        file = await ctx.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        b64 = base64.b64encode(pdf_bytes).decode()

        await msg.edit_text(f"🤖 Надсилаю до Claude ({len(pdf_bytes)//1024} KB)...")

        result = await ask_claude(b64)
        rows = result.get("rows", [])
        ved  = result.get("vedomist_type", "?")
        dil  = result.get("dilinitsa", "?")
        per  = result.get("period", "?")

        if not rows:
            await msg.edit_text("⚠️ Рядків не знайдено")
            return

        header = f"📋 Відомість тип {ved} · Діл. {dil}\n📅 {per}\n👥 {len(rows)} осіб\n{'─'*30}\n"
        chunks = []
        current = header

        for i, r in enumerate(rows, 1):
            name     = r.get("name","?")
            addr     = r.get("address","?").replace("Сошичне, ","").replace(";;","").strip()
            summa    = fmt_hrn(r.get("sum", 0))
            date     = r.get("pay_date","?")
            passport = r.get("passport","")
            line = f"{i}. {name}\n   📍 {addr}\n   💰 {summa} · 📅 {date}\n   🪪 {passport}\n\n"

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

        total = sum(float(r.get("sum",0)) for r in rows)
        await update.message.reply_text(
            f"✅ Готово!\n💰 Загальна сума: {fmt_hrn(total)}\n📄 {doc.file_name}"
        )

    except Exception as e:
        await msg.edit_text(f"❌ Помилка: {e}")

async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📄 Надішли PDF файл відомості")

async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot started")
    await app.run_polling()

if __name__ == "__main__":
    asyncio.run(main())
