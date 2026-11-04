import os
import logging
import sqlite3
import asyncio
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
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
        time.sleep(300)
        try:
            app_url = os.environ.get("RENDER_EXTERNAL_URL") or f"http://127.0.0.1:{os.environ.get('PORT', 8080)}"
            urllib.request.urlopen(app_url, timeout=20)
        except Exception:
            pass

threading.Thread(target=self_ping, daemon=True).start()

BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

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
        'welcome': "👋 Assalomu alaykum! Botimizga xush kelibsiz. Kerakli xizmatni tanlang[span_4](start_span)[span_4](end_span):",
        'img_to_pdf': "🖼 Rasm ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Rasm",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 PDF birlashtirish[span_5](start_span)[span_5](end_span)",
        'edit_image': "✨ Rasmni tahrirlash (Resize/Piksel)",
        'back': "⬅️ Orqaga",
        'send_img': "📥 Rasmlarni yuboring. Tugatgach '✅ Tayyor'ni bosing[span_6](start_span)[span_6](end_span).",
        'ready': "✅ Tayyor",
        'cancel': "❌ Bekor qilish",
        'resize_info': "📐 Rasm hajmini (piksellarini) o‘zgartirish uchun rasmni yuboring va yangi o'lchamni yuboring (masalan: 800x600)."
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
        'edit_image': "✨ Расмни таҳрирлаш (Resize)",
        'back': "⬅️ Орқага",
        'send_img': "📥 Расмларни юборинг. Тугатгач '✅ Тайёр'ни босинг.",
        'ready': "✅ Тайёр",
        'cancel': "❌ Бекор қилиш",
        'resize_info': "📐 Расм ҳажмини ўзгартириш учун расмни юборинг."
    },
    'ru': {
        'choose_lang': "🌐 Пожалуйста, выберите язык:",
        'welcome': "👋 Добро пожаловать! Выберите нужную услугу[span_7](start_span)[span_7](end_span):",
        'img_to_pdf': "🖼 Картинка ➡️ PDF",
        'pdf_to_img': "📄 PDF ➡️ Картинка",
        'word_to_pdf': "📝 WORD ➡️ PDF",
        'pdf_to_word': "📁 PDF ➡️ WORD",
        'excel_to_pdf': "📊 EXCEL ➡️ PDF",
        'pdf_merge': "🔗 Объединить PDF",
        'edit_image': "✨ Редактировать фото (Resize)",
        'back': "⬅️ Назад",
        'send_img': "📥 Отправьте изображения. Нажмите '✅ Готово', когда закончите[span_8](start_span)[span_8](end_span).",
        'ready': "✅ Готово",
        'cancel': "❌ Отмена",
        'resize_info': "📐 Отправьте фото для изменения размера."
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
        'edit_image': "✨ Edit Image (Resize)",
        'back': "⬅️ Back",
        'send_img': "📥 Send images. Click '✅ Done' when finished.",
        'ready': "✅ Done",
        'cancel': "❌ Cancel",
        'resize_info': "📐 Send image to resize."
    }
}

def get_text(user_id, key):
    lang = get_user_lang(user_id) or 'uz_lat'
    return TRANSLATIONS.get(lang, TRANSLATIONS['uz_lat']).get(key, "")

# --- KLAVIATURALAR ---
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

# --- /START VA TIL TANLASH ---
@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    user_id = message.from_user.id
    lang = get_user_lang(user_id)
    if not lang:
        await message.answer("🌐 Iltimos, tilni tanlang / Пожалуйста, выберите язык:", reply_markup=lang_keyboard())
    else:
        await message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))

@dp.callback_query(F.data.startswith("lang:"))
async def process_lang(callback: types.CallbackQuery):
    lang = callback.data.split(":")[1]
    user_id = callback.from_user.id
    set_user_lang(user_id, lang, callback.from_user.full_name)
    await callback.message.edit_text("✅ Til o'zgartirildi!")
    await callback.message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))
    await callback.answer()

# --- MENYU TUGMALARI BO'YICHA AMALLAR ---
@dp.message(F.text.in_({
    TRANSLATIONS['uz_lat']['img_to_pdf'], TRANSLATIONS['uz_cyr']['img_to_pdf'], 
    TRANSLATIONS['ru']['img_to_pdf'], TRANSLATIONS['en']['img_to_pdf']
}))
async def img_to_pdf_menu(message: types.Message):
    user_id = message.from_user.id
    builder = ReplyKeyboardBuilder()
    builder.button(text=get_text(user_id, 'ready'))
    builder.button(text=get_text(user_id, 'cancel'))
    builder.adjust(2)
    await message.answer(get_text(user_id, 'send_img'), reply_markup=builder.as_markup(resize_keyboard=True))

@dp.message(F.text.in_({
    TRANSLATIONS['uz_lat']['edit_image'], TRANSLATIONS['uz_cyr']['edit_image'], 
    TRANSLATIONS['ru']['edit_image'], TRANSLATIONS['en']['edit_image']
}))
async def edit_image_menu(message: types.Message):
    user_id = message.from_user.id
    await message.answer(get_text(user_id, 'resize_info'))

@dp.message(F.text.in_({
    TRANSLATIONS['uz_lat']['cancel'], TRANSLATIONS['uz_cyr']['cancel'], 
    TRANSLATIONS['ru']['cancel'], TRANSLATIONS['en']['cancel']
}))
async def cancel_action(message: types.Message):
    user_id = message.from_user.id
    await message.answer(get_text(user_id, 'welcome'), reply_markup=main_menu_keyboard(user_id))

# --- RASMNI PIKSEL VA RAZMERINI O'ZGARTIRISH (RESIZE) ---
@dp.message(F.photo)
async def handle_photo(message: types.Message):
    user_id = message.from_user.id
    photo = message.photo[-1]
    
    file = await bot.get_file(photo.file_id)
    os.makedirs("downloads", exist_ok=True)
    file_path = f"downloads/{user_id}_{photo.file_unique_id}.jpg"
    await bot.download_file(file.file_path, file_path)

    # Odatiy holatda rasm o'lchamini ikki barobar kichraytirish yoki tahrirlash namunasi
    img = Image.open(file_path)
    width, height = img.size
    
    # Piksellarini va o'lchamini to'g'irlab, foydalanuvchiga yuborish
    resized_path = f"downloads/resized_{user_id}.jpg"
    img.resize((width // 2, height // 2)).save(resized_path)

    await message.answer_photo(
        types.FSInputFile(resized_path),
        caption=f"✅ Rasm piksellari moslashtirildi!\nAsl o'lcham: {width}x{height}\nYangi o'lcham: {width//2}x{height//2}"
    )
    
    os.remove(file_path)
    os.remove(resized_path)

async def main():
    await start_web_server()
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main))
