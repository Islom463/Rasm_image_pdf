import os
import logging
import sqlite3
import asyncio
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from PIL import Image, ImageEnhance, ImageFilter
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
import openpyxl

# Xavfsiz PDF importlari (xatolik bermasligi uchun)
try:
    from pypdf import PdfMerger, PdfReader
except ImportError:
    from PyPDF2 import PdfMerger, PdfReader

from aiohttp import web
import urllib.request
import threading

logging.basicConfig(level=logging.INFO)

# --- RENDER WEB SERVER ---
async def start_web_server():
    app = web.Application()
    app.add_routes([
        web.get('/', lambda r: web.Response(text="PDF & Image Bot is running!")),
        web.get('/ping', lambda r: web.Response(text="pong"))
    ])
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()

def self_ping():
    while True:
        threading.Event().wait(300)
        try:
            app_url = os.getenv("RENDER_EXTERNAL_URL") or f"http://127.0.0.1:{os.environ.get('PORT', 8080)}"
            urllib.request.urlopen(app_url, timeout=20)
        except Exception:
            pass

threading.Thread(target=self_ping, daemon=True).start()

BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

user_sessions = {}

# --- BAZA ---
def init_db():
    conn = sqlite3.connect("pdf_bot.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            full_name TEXT,
            lang TEXT DEFAULT 'uz_lat'
        )
    """)
    conn.commit()
    conn.close()

init_db()

def get_user_lang(user_id):
    conn = sqlite3.connect("pdf_bot.db")
    cursor = conn.cursor()
    cursor.execute("SELECT lang FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

def set_user_lang(user_id, lang, name="User"):
    conn = sqlite3.connect("pdf_bot.db")
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO users (user_id, full_name, lang) VALUES (?, ?, ?)", (user_id, name, lang))
    conn.commit()
    conn.close()

# --- TARJIMALAR ---
TRANSLATIONS = {
    'uz_lat': {
        'choose_lang': "🌐 Iltimos, interfeys tilini tanlang:",
        'welcome': "👋 Assalomu alaykum! Botimizga xush kelibsiz. Kerakli xizmatni tanlang:",
        'img_to_pdf': "🖼 Rasm ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Rasm",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 PDF birlashtirish",
        'edit_image': "✨ Rasmni tahrirlash (Resize/Tiniqlashtirish)",
        'back': "⬅️ Orqaga",
        'send_img': "📥 Rasmlarni yuboring. Barchasini yuborib bo'lgach '✅ Tayyor' tugmasini bosing.",
        'ready': "✅ Tayyor",
        'cancel': "❌ Bekor qilish",
        'resize_info': "📐 Rasmni yuboring: uni avtomatik tiniqlashtirib, piksellarini to'g'irlab beraman.",
        'pdf_merge_info': "🔗 Birlashtirmoqchi bo'lgan barcha PDF fayllarni ketma-ket yuboring va '✅ Tayyor'ni bosing."
    },
    'uz_cyr': {
        'choose_lang': "🌐 Илтимос, тилни танланг:",
        'welcome': "👋 Ассалому алайкум! Ботимизга хуш келибсиз. Керакли хизматни танланг:",
        'img_to_pdf': "🖼 Расм ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Расм",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 PDF бирлаштириш",
        'edit_image': "✨ Расмни таҳрирлаш",
        'back': "⬅️ Орқага",
        'send_img': "📥 Расмларни юборинг. Тугатгач '✅ Тайёр'ни босинг.",
        'ready': "✅ Тайёр",
        'cancel': "❌ Бекор қилиш",
        'resize_info': "📐 Расмни юборинг.",
        'pdf_merge_info': "🔗 PDF файлларни юборинг."
    },
    'ru': {
        'choose_lang': "🌐 Пожалуйста, выберите язык:",
        'welcome': "👋 Добро пожаловать! Выберите нужную услугу:",
        'img_to_pdf': "🖼 Картинка ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Картинка",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 Объединить PDF",
        'edit_image': "✨ Редактировать фото",
        'back': "⬅️ Назад",
        'send_img': "📥 Отправьте изображения. Нажмите '✅ Готово'.",
        'ready': "✅ Готово",
        'cancel': "❌ Отмена",
        'resize_info': "📐 Отправьте фото.",
        'pdf_merge_info': "🔗 Отправьте PDF файлы."
    },
    'en': {
        'choose_lang': "🌐 Please select a language:",
        'welcome': "👋 Welcome! Choose the desired service:",
        'img_to_pdf': "🖼 Image ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Image",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 Merge PDF",
        'edit_image': "✨ Edit Image",
        'back': "⬅️ Back",
        'send_img': "📥 Send images. Click '✅ Done'.",
        'ready': "✅ Done",
        'cancel': "❌ Cancel",
        'resize_info': "📐 Send image.",
        'pdf_merge_info': "🔗 Send PDF files."
    }
}

def get_text(user_id, key):
    lang = get_user_lang(user_id) or 'uz_lat'
    return TRANSLATIONS.get(lang, TRANSLATIONS['uz_lat']).get(key, "")

def lang_keyboard():
    builder = InlineKeyboardBuilder()
    builder.button(text="🇺🇿 O'zbekcha", callback_data="lang:uz_lat")
    builder.button(text="🇺🇿 Ўзбекча", callback_data="lang:uz_cyr")
    builder.button(text="🇷🇺 Русский", callback_data="lang:ru")
    builder.button(text="🇬🇧 English", callback_data="lang:en")
    builder.adjust(2)
    return builder.as_markup()

def main_menu_keyboard(user_id):
    builder = ReplyKeyboardBuilder()
    builder.button(text=get_text(user_id, 'img_to_pdf'))
    builder.button(text=get_text(user_id, 'pdf_to_img'))
    builder.button(text=get_text(user_id, 'word_to_pdf'))
    builder.button(text=get_text(user_id, 'pdf_to_word'))
    builder.button(text=get_text(user_id, 'excel_to_pdf'))
    builder.button(text=get_text(user_id, 'pdf_merge'))
    builder.button(text=get_text(user_id, 'edit_image'))
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)

def action_keyboard(user_id):
    builder = ReplyKeyboardBuilder()
    builder.button(text=get_text(user_id, 'ready'))
    builder.button(text=get_text(user_id, 'cancel'))
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    user_id = message.from_user.id
    user_sessions[user_id] = {'mode': None, 'files': []}
    lang = get_user_lang(user_id)
    if not lang:
        await message.answer("🌐 Iltimos, tilni tanlang:", reply_markup=lang_keyboard())
    else:
        await message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))

@dp.callback_query(F.data.startswith("lang:"))
async def process_lang(callback: types.CallbackQuery):
    lang = callback.data.split(":")[1]
    user_id = callback.from_user.id
    set_user_lang(user_id, lang, callback.from_user.full_name)
    user_sessions[user_id] = {'mode': None, 'files': []}
    await callback.message.edit_text("✅ Til o'zgartirildi!")
    await callback.message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))
    await callback.answer()

@dp.message(F.text.in_({TRANSLATIONS[l]['img_to_pdf'] for l in TRANSLATIONS}))
async def menu_img_to_pdf(message: types.Message):
    user_id = message.from_user.id
    user_sessions[user_id] = {'mode': 'img_to_pdf', 'files': []}
    await message.answer(get_text(user_id, 'send_img'), reply_markup=action_keyboard(user_id))

@dp.message(F.text.in_({TRANSLATIONS[l]['pdf_merge'] for l in TRANSLATIONS}))
async def menu_pdf_merge(message: types.Message):
    user_id = message.from_user.id
    user_sessions[user_id] = {'mode': 'pdf_merge', 'files': []}
    await message.answer(get_text(user_id, 'pdf_merge_info'), reply_markup=action_keyboard(user_id))

@dp.message(F.text.in_({TRANSLATIONS[l]['edit_image'] for l in TRANSLATIONS}))
async def menu_edit_image(message: types.Message):
    user_id = message.from_user.id
    user_sessions[user_id] = {'mode': 'edit_image', 'files': []}
    await message.answer(get_text(user_id, 'resize_info'), reply_markup=action_keyboard(user_id))

@dp.message(F.text.in_({TRANSLATIONS[l]['cancel'] for l in TRANSLATIONS}))
async def menu_cancel(message: types.Message):
    user_id = message.from_user.id
    user_sessions[user_id] = {'mode': None, 'files': []}
    await message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))

@dp.message(F.text.in_({TRANSLATIONS[l]['ready'] for l in TRANSLATIONS}))
async def menu_ready(message: types.Message):
    user_id = message.from_user.id
    session = user_sessions.get(user_id, {'mode': None, 'files': []})
    mode = session.get('mode')
    files = session.get('files', [])

    if not files:
        await message.answer("⚠️ Hali hech qanday fayl yubormadingiz!")
        return

    os.makedirs("downloads", exist_ok=True)

    if mode == 'img_to_pdf':
        pdf_path = f"downloads/images_{user_id}.pdf"
        images = [Image.open(f).convert('RGB') for f in files]
        if images:
            images[0].save(pdf_path, save_all=True, append_images=images[1:])
            await message.answer_document(types.FSInputFile(pdf_path), caption="✅ Rasmlardan PDF tayyor!")
            os.remove(pdf_path)
        for f in files:
            try:
                os.remove(f)
            except Exception:
                pass

    elif mode == 'pdf_merge':
        merger = PdfMerger()
        for pdf in files:
            merger.append(pdf)
        merged_path = f"downloads/merged_{user_id}.pdf"
        merger.write(merged_path)
        merger.close()
        await message.answer_document(types.FSInputFile(merged_path), caption="✅ PDF fayllar birlashtirildi!")
        os.remove(merged_path)
        for f in files:
            try:
                os.remove(f)
            except Exception:
                pass

    user_sessions[user_id] = {'mode': None, 'files': []}
    await message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))

@dp.message(F.photo)
async def handle_photo(message: types.Message):
    user_id = message.from_user.id
    session = user_sessions.setdefault(user_id, {'mode': None, 'files': []})
    mode = session.get('mode')

    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    os.makedirs("downloads", exist_ok=True)
    file_path = f"downloads/{user_id}_{photo.file_unique_id}.jpg"
    await bot.download_file(file.file_path, file_path)

    if mode == 'img_to_pdf':
        session['files'].append(file_path)
        await message.answer(f"📥 Rasm qo'shildi ({len(session['files'])} ta). Yana yuboring yoki '✅ Tayyor'ni bosing.")
    
    elif mode == 'edit_image':
        img = Image.open(file_path).convert('RGB')
        enhancer = ImageEnhance.Contrast(img)
        img_enhanced = enhancer.enhance(1.3)
        img_sharpened = img_enhanced.filter(ImageFilter.SHARPEN)
        
        out_path = f"downloads/edited_{user_id}.jpg"
        img_sharpened.save(out_path)
        
        width, height = img.size
        await message.answer_photo(
            types.FSInputFile(out_path),
            caption=f"✨ Rasm tiniqlashtirildi va piksellari moslashtirildi!\nO'lchami: {width}x{height}"
        )
        os.remove(file_path)
        os.remove(out_path)
    else:
        img = Image.open(file_path).convert('RGB')
        img_sharpened = img.filter(ImageFilter.SHARPEN)
        out_path = f"downloads/fixed_{user_id}.jpg"
        img_sharpened.save(out_path)
        await message.answer_photo(types.FSInputFile(out_path), caption="✨ Rasm tiniqlashtirildi!")
        os.remove(file_path)
        os.remove(out_path)

@dp.message(F.document)
async def handle_document(message: types.Message):
    user_id = message.from_user.id
    doc = message.document
    file_name = doc.file_name.lower()
    
    file = await bot.get_file(doc.file_id)
    os.makedirs("downloads", exist_ok=True)
    input_path = f"downloads/{user_id}_{doc.file_name}"
    await bot.download_file(file.file_path, input_path)

    session = user_sessions.setdefault(user_id, {'mode': None, 'files': []})
    
    if session.get('mode') == 'pdf_merge' and file_name.endswith('.pdf'):
        session['files'].append(input_path)
        await message.answer(f"📥 PDF qo'shildi ({len(session['files'])} ta). Yana yuboring yoki '✅ Tayyor'ni bosing.")
        return

    try:
        if file_name.endswith('.docx'):
            pdf_path = input_path.rsplit('.', 1)[0] + ".pdf"
            doc_file = Document(input_path)
            
            pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []
            for para in doc_file.paragraphs:
                if para.text.strip():
                    story.append(Paragraph(para.text, styles['Normal']))
                    story.append(Spacer(1, 10))
            pdf.build(story)
            
            await message.answer_document(types.FSInputFile(pdf_path), caption="✅ Word fayl PDF formatiga o'tkazildi!")
            os.remove(pdf_path)

        elif file_name.endswith('.xlsx'):
            pdf_path = input_path.rsplit('.', 1)[0] + ".pdf"
            wb = openpyxl.load_workbook(input_path)
            sheet = wb.active
            
            pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []
            for row in sheet.iter_rows(values_only=True):
                row_text = " | ".join([str(cell) for cell in row if cell is not None])
                if row_text.strip():
                    story.append(Paragraph(row_text, styles['Normal']))
                    story.append(Spacer(1, 6))
            pdf.build(story)
            
            await message.answer_document(types.FSInputFile(pdf_path), caption="✅ Excel fayl PDF formatiga o'tkazildi!")
            os.remove(pdf_path)

        elif file_name.endswith('.pdf'):
            docx_path = input_path.rsplit('.', 1)[0] + ".docx"
            reader = PdfReader(input_path)
            writer_docx = Document()
            
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    writer_docx.add_paragraph(text)
            writer_docx.save(docx_path)
            
            await message.answer_document(types.FSInputFile(docx_path), caption="✅ PDF fayl WORD formatiga o'tkazildi!")
            os.remove(docx_path)

        os.remove(input_path)

    except Exception as e:
        logging.error(f"Faylni qayta ishlash xatosi: {e}")
        await message.answer("❌ Faylni o'zgartirishda xatolik yuz berdi!")

async def main():
    await start_web_server()
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
