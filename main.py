"""
PDF + Rasm + AI hujjatlar boti.

Imkoniyatlar:
  Fayllar: Rasm->PDF, PDF->Rasm, Word->PDF, PDF->Word, Excel->PDF, PDF birlashtirish, rasmni tiniqlashtirish
  AI (Groq): Taqdimot (PPTX), Referat/Mustaqil ish, Kurs ishi, Dars ishlanma, Rezyume/Obektivka,
             Tadqiqot (ilmiy ish), Boshqa ishlar
  Qo'lda yozish: matnni daftar varag'iga qo'lyozma ko'rinishida chiqaradi

Render Environment:
  BOT_TOKEN       - majburiy
  GROQ_API_KEY    - AI bo'limlari uchun (gsk_...)
  ADMIN_ID        - ixtiyoriy (AI limitidan ozod qilish)
  AI_DAILY_LIMIT  - ixtiyoriy, standart 10
"""
import os
import re
import json
import html
import time
import uuid
import random
import shutil
import asyncio
import logging
import sqlite3
import threading
import urllib.request
from contextlib import closing
from datetime import datetime

import aiohttp
from aiohttp import web
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from PIL import Image, ImageEnhance, ImageFilter, ImageDraw, ImageFont, ImageOps

logging.basicConfig(level=logging.INFO)

# ---------------- Sozlamalar ----------------
DB_PATH = os.getenv("DB_PATH", "pdf_bot.db")
TMP_DIR = "downloads"
FONT_DIR = "fonts"
MAX_DOWNLOAD = 20 * 1024 * 1024  # Telegram Bot API yuklab olish chegarasi
ADMIN_IDS = {int(x) for x in re.findall(r"\d+", os.getenv("ADMIN_ID", "") + " " + os.getenv("ADMIN_IDS", ""))}
try:
    AI_DAILY_LIMIT = int(os.getenv("AI_DAILY_LIMIT", "10"))
except ValueError:
    AI_DAILY_LIMIT = 10

os.makedirs(TMP_DIR, exist_ok=True)
os.makedirs(FONT_DIR, exist_ok=True)


# ---------------- Render uchun veb-server va keep-alive ----------------
async def _root_handler(request):
    return web.Response(text="PDF & Image Bot is running!")


async def _ping_handler(request):
    return web.Response(text="pong")


async def start_web_server():
    app = web.Application()
    app.add_routes([web.get("/", _root_handler), web.get("/ping", _ping_handler)])
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 8080))
    await web.TCPSite(runner, "0.0.0.0", port).start()


def self_ping():
    while True:
        time.sleep(240)
        try:
            url = os.getenv("RENDER_EXTERNAL_URL") or f"http://127.0.0.1:{os.environ.get('PORT', 8080)}"
            urllib.request.urlopen(url, timeout=20).read()
        except Exception:
            pass


threading.Thread(target=self_ping, daemon=True).start()


# ---------------- Kalitlar ----------------
def _clean(v):
    return (v or "").strip().strip('"').strip("'").strip()


def detect_keys():
    env = {k: _clean(v) for k, v in os.environ.items()}
    tok_re = r"\d{6,}:[A-Za-z0-9_-]{30,}"
    token = ""
    for k in ("BOT_TOKEN", "TELEGRAM_BOT_TOKEN", "TOKEN"):
        if re.fullmatch(tok_re, env.get(k, "")):
            token = env[k]
            break
    if not token:
        for v in env.values():
            if re.fullmatch(tok_re, v):
                token = v
                break
    groq = next((v for v in env.values() if v.startswith("gsk_")), "")
    return token, groq


BOT_TOKEN, KEY_GROQ = detect_keys()
logging.info("Kalitlar: BOT=%s GROQ=%s", bool(BOT_TOKEN), bool(KEY_GROQ))


def stop_with_error(msg):
    while True:
        logging.error(msg)
        time.sleep(30)


if not BOT_TOKEN:
    stop_with_error("BOT_TOKEN topilmadi! Render Environment'ga BOT_TOKEN ni kiriting.")

bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

user_sessions = {}
HEAVY = asyncio.Semaphore(2)


# ---------------- Lotin -> Kirill (o'zbekcha) ----------------
PROTECT = {"PDF", "Word", "WORD", "Excel", "EXCEL", "PowerPoint", "PPTX", "DOCX", "AI", "Telegram", "Groq"}
_SINGLE = {
    "a": "а", "b": "б", "d": "д", "f": "ф", "g": "г", "h": "ҳ", "i": "и", "j": "ж", "k": "к", "l": "л",
    "m": "м", "n": "н", "o": "о", "p": "п", "q": "қ", "r": "р", "s": "с", "t": "т", "u": "у", "v": "в",
    "x": "х", "y": "й", "z": "з", "c": "с", "w": "в",
}
_WORD_RE = re.compile(
    r"(\{[^}]*\}|<[^>]*>|https?://\S+|@\w+|(?<![\w])/\w+|`[^`]*`)|([A-Za-z][A-Za-z'’ʻʼ‘`]*)"
)


def _conv_word(w):
    w = re.sub(r"[’ʻʼ‘`]", "'", w)
    low = w.lower()
    out, i, n = [], 0, len(w)
    while i < n:
        c, two, up = low[i], low[i:i + 2], w[i].isupper()

        def put(s):
            out.append(s.upper() if up else s)

        if two == "o'":
            put("ў"); i += 2; continue
        if two == "g'":
            put("ғ"); i += 2; continue
        if two == "sh":
            put("ш"); i += 2; continue
        if two == "ch":
            put("ч"); i += 2; continue
        if two == "yo":
            put("ё"); i += 2; continue
        if two == "yu":
            put("ю"); i += 2; continue
        if two == "ya":
            put("я"); i += 2; continue
        if two == "ye":
            put("е"); i += 2; continue
        if c == "e":
            put("э" if i == 0 else "е"); i += 1; continue
        if c == "'":
            if 0 < i < n - 1:
                out.append("ъ")
            i += 1; continue
        put(_SINGLE.get(c, w[i]))
        i += 1
    return "".join(out)


def lat2cyr(text):
    def repl(m):
        if m.group(1):
            return m.group(1)
        w = m.group(2)
        if w in PROTECT or (len(w) > 1 and w.isupper()):
            return w
        return _conv_word(w)

    return _WORD_RE.sub(repl, text)


# ---------------- Matnlar ----------------
LANGS = ("uz_lat", "uz_cyr", "ru", "en")
LANG_NAMES = {
    "uz_lat": "Uzbek (Latin script)", "uz_cyr": "Uzbek (Cyrillic script)",
    "ru": "Russian", "en": "English",
}

UZ = {
    "choose_lang": "🌐 Iltimos, interfeys tilini tanlang:",
    "welcome": "👋 Assalomu alaykum! Men fayllarni o'zgartiraman va AI yordamida slayd, referat, kurs ishi, rezyume kabi ishlarni tayyorlab beraman.\n\nBo'limni tanlang 👇",
    "m_files": "📁 Fayllarni o'zgartirish",
    "m_ai": "🤖 AI bilan tayyorlash",
    "handwrite": "✍️ Qo'lda yozish",
    "edit_image": "✨ Rasmni tiniqlashtirish",
    "help": "📖 Qo'llanma",
    "img_to_pdf": "🖼 Rasm ➡️ PDF",
    "pdf_to_img": "📄 PDF ➡️ Rasm",
    "word_to_pdf": "📝 WORD ➡️ PDF",
    "pdf_to_word": "📁 PDF ➡️ WORD",
    "excel_to_pdf": "📊 EXCEL ➡️ PDF",
    "pdf_merge": "🔗 PDF birlashtirish",
    "ai_slides": "📕 Taqdimot (slayd)",
    "ai_referat": "📘 Referat / Mustaqil ish",
    "ai_kurs": "📚 Kurs ishi",
    "ai_dars": "📒 Dars ishlanma",
    "ai_resume": "📄 Rezyume / Obektivka",
    "ai_research": "🔬 Tadqiqot (ilmiy ish)",
    "ai_other": "🧾 Boshqa ishlar",
    "back": "⬅️ Orqaga",
    "ready": "✅ Tayyor",
    "cancel": "❌ Bekor qilish",
    "files_menu": "📁 Qaysi amalni bajaramiz?",
    "ai_menu": "🤖 Qaysi ishni tayyorlaymiz?",
    "p_img_to_pdf": "📥 Rasmlarni yuboring (rasm yoki fayl ko'rinishida). Hammasini yuborib bo'lgach «✅ Tayyor» tugmasini bosing.",
    "p_pdf_merge": "🔗 Birlashtiriladigan PDF fayllarni ketma-ket yuboring va «✅ Tayyor» ni bosing.",
    "p_pdf_to_img": "📄 PDF faylni yuboring — har bir sahifasini rasmga aylantiraman.",
    "p_word_to_pdf": "📝 Word (.docx) faylni yuboring.",
    "p_pdf_to_word": "📁 PDF faylni yuboring — matnini Word (.docx) ga o'tkazaman.",
    "p_excel_to_pdf": "📊 Excel (.xlsx) faylni yuboring.",
    "p_edit_image": "✨ Rasmni yuboring — tiniqlashtirib, kontrastini yaxshilab beraman.\n💡 Eng yaxshi natija uchun rasmni «fayl» sifatida yuboring.",
    "p_handwrite": "✍️ Matnni yozib yuboring (yoki .txt fayl) — daftar varag'iga qo'lyozma qilib beraman.",
    "p_ai_slides": "📕 <b>Taqdimot</b> mavzusini yozing.\nMasalan: <i>Sun'iy intellekt va kelajak, 12 ta slayd</i>",
    "p_ai_referat": "📘 <b>Referat / Mustaqil ish</b> mavzusini yozing.\nMasalan: <i>Global isish muammosi, 10 bet</i>",
    "p_ai_kurs": "📚 <b>Kurs ishi</b> mavzusi va fanini yozing.\nMasalan: <i>Marketing: Raqamli marketing strategiyalari, 20 bet</i>",
    "p_ai_dars": "📒 <b>Dars ishlanma</b>: fan, sinf va mavzuni yozing.\nMasalan: <i>Matematika, 5-sinf, Oddiy kasrlar</i>",
    "p_ai_resume": "📄 <b>Rezyume / Obektivka</b>: o'zingiz haqingizda yozing — ism-familiya, tug'ilgan yil, ta'lim, ish tajribasi, ko'nikmalar, tillar, telefon.\n«Obektivka» kerak bo'lsa shuni yozing.",
    "p_ai_research": "🔬 <b>Tadqiqot (ilmiy ish)</b> mavzusini yozing.\nMasalan: <i>Talabalarda onlayn ta'limning samaradorligi, 8 bet</i>",
    "p_ai_other": "🧾 Qanday ish kerak? Yozing: esse, maqola, tezis, ariza, xat, test savollari, ssenariy...\nMasalan: <i>«Vatan» mavzusida esse</i>",
    "no_files": "⚠️ Hali hech narsa yubormadingiz!",
    "working": "⏳ Ishlanmoqda...",
    "ai_wait": "⏳ Tayyorlanmoqda, 30–90 soniya kuting...",
    "ai_note": "⚠️ AI yozgan matn: faktlar va manbalarni albatta tekshirib oling.",
    "error": "❌ Xatolik yuz berdi. Qayta urinib ko'ring.",
    "too_big": "❌ Fayl juda katta (20 MB dan oshmasin).",
    "wrong_type": "⚠️ Bu rejimga mos fayl emas. Kerakli fayl turi: {need}",
    "limit": "⏳ Bugungi AI limiti ({n} ta) tugadi. Ertaga qayta urinib ko'ring.",
    "no_ai": "⚠️ AI ishlamayapti: GROQ_API_KEY topilmadi.",
    "pick_pdf": "📄 PDF bilan nima qilamiz?",
    "send_text": "✍️ Matn yuboring.",
    "menu_hint": "Bo'limni menyudan tanlang 👇",
    "help_text": (
        "📖 <b>Qo'llanma</b>\n\n"
        "📁 <b>Fayllar</b>: rasm → PDF, PDF → rasm, Word/Excel → PDF, PDF → Word, PDF birlashtirish.\n"
        "Rasm yuborsangiz — PDF yig'ishni boshlayman. Word/Excel fayl yuborsangiz — PDF qilib beraman.\n\n"
        "🤖 <b>AI bilan</b>: mavzuni yozasiz — tayyor ish keladi.\n"
        "• Taqdimot → .pptx (PowerPoint)\n• Referat, kurs ishi, tadqiqot, dars ishlanma, rezyume → .docx va .pdf\n"
        "• Hajmni yozing: <i>10 bet</i> yoki <i>12 ta slayd</i>\n\n"
        "✍️ <b>Qo'lda yozish</b>: matn → daftar varag'idagi qo'lyozma.\n"
        "✨ <b>Tiniqlashtirish</b>: xira rasmni yaxshilaydi.\n\n"
        "ℹ️ Word → PDF da rasmlar ko'chirilmaydi; PDF → Word faqat matnni oladi (skaner qilingan PDF uchun ishlamaydi)."
    ),
}

RU = {
    "choose_lang": "🌐 Пожалуйста, выберите язык:",
    "welcome": "👋 Здравствуйте! Я конвертирую файлы и с помощью ИИ готовлю презентации, рефераты, курсовые, резюме.\n\nВыберите раздел 👇",
    "m_files": "📁 Конвертация файлов", "m_ai": "🤖 Подготовка с ИИ", "handwrite": "✍️ Рукописный текст",
    "edit_image": "✨ Улучшить фото", "help": "📖 Справка",
    "img_to_pdf": "🖼 Фото ➡️ PDF", "pdf_to_img": "📄 PDF ➡️ Фото", "word_to_pdf": "📝 WORD ➡️ PDF",
    "pdf_to_word": "📁 PDF ➡️ WORD", "excel_to_pdf": "📊 EXCEL ➡️ PDF", "pdf_merge": "🔗 Объединить PDF",
    "ai_slides": "📕 Презентация", "ai_referat": "📘 Реферат / Самостоятельная", "ai_kurs": "📚 Курсовая работа",
    "ai_dars": "📒 Конспект урока", "ai_resume": "📄 Резюме / Анкета", "ai_research": "🔬 Исследование",
    "ai_other": "🧾 Другие работы", "back": "⬅️ Назад", "ready": "✅ Готово", "cancel": "❌ Отмена",
    "working": "⏳ Обработка...", "ai_wait": "⏳ Готовлю, подождите 30–90 секунд...",
    "error": "❌ Произошла ошибка. Попробуйте ещё раз.",
    "menu_hint": "Выберите раздел в меню 👇",
}

EN = {
    "choose_lang": "🌐 Please select a language:",
    "welcome": "👋 Hello! I convert files and use AI to prepare presentations, essays, coursework and résumés.\n\nChoose a section 👇",
    "m_files": "📁 File conversion", "m_ai": "🤖 AI documents", "handwrite": "✍️ Handwriting",
    "edit_image": "✨ Enhance image", "help": "📖 Help",
    "img_to_pdf": "🖼 Image ➡️ PDF", "pdf_to_img": "📄 PDF ➡️ Image", "word_to_pdf": "📝 WORD ➡️ PDF",
    "pdf_to_word": "📁 PDF ➡️ WORD", "excel_to_pdf": "📊 EXCEL ➡️ PDF", "pdf_merge": "🔗 Merge PDF",
    "ai_slides": "📕 Presentation", "ai_referat": "📘 Essay / Report", "ai_kurs": "📚 Coursework",
    "ai_dars": "📒 Lesson plan", "ai_resume": "📄 Résumé / CV", "ai_research": "🔬 Research paper",
    "ai_other": "🧾 Other documents", "back": "⬅️ Back", "ready": "✅ Done", "cancel": "❌ Cancel",
    "working": "⏳ Working...", "ai_wait": "⏳ Preparing, please wait 30–90 seconds...",
    "error": "❌ An error occurred. Please try again.",
    "menu_hint": "Choose a section from the menu 👇",
}


def tr_lang(lang, key):
    if lang == "uz_cyr":
        return lat2cyr(UZ.get(key, key))
    table = {"ru": RU, "en": EN}.get(lang)
    if table and key in table:
        return table[key]
    return UZ.get(key, key)


BUTTON_KEYS = [
    "m_files", "m_ai", "handwrite", "edit_image", "help",
    "img_to_pdf", "pdf_to_img", "word_to_pdf", "pdf_to_word", "excel_to_pdf", "pdf_merge",
    "ai_slides", "ai_referat", "ai_kurs", "ai_dars", "ai_resume", "ai_research", "ai_other",
    "back", "ready", "cancel",
]
LABEL_TO_KEY = {}
for _l in LANGS:
    for _k in BUTTON_KEYS:
        LABEL_TO_KEY.setdefault(tr_lang(_l, _k), _k)
ALL_LABELS = set(LABEL_TO_KEY)

FILE_MODES = {"img_to_pdf", "pdf_to_img", "word_to_pdf", "pdf_to_word", "excel_to_pdf", "pdf_merge", "edit_image", "handwrite"}
AI_MODES = {"ai_slides", "ai_referat", "ai_kurs", "ai_dars", "ai_resume", "ai_research", "ai_other"}

MAIN_ROWS = [["m_files", "m_ai"], ["handwrite", "edit_image"], ["help"]]
FILES_ROWS = [["img_to_pdf", "pdf_to_img"], ["word_to_pdf", "pdf_to_word"], ["excel_to_pdf", "pdf_merge"], ["back"]]
AI_ROWS = [["ai_slides", "ai_referat"], ["ai_kurs", "ai_dars"], ["ai_resume", "ai_research"], ["ai_other"], ["back"]]
ACTION_ROWS = [["ready", "cancel"]]
CANCEL_ROWS = [["cancel"]]


# ---------------- Baza (til) ----------------
def init_db():
    with closing(sqlite3.connect(DB_PATH)) as c:
        c.execute("CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY, full_name TEXT, lang TEXT DEFAULT 'uz_lat')")
        c.commit()


init_db()


def get_user_lang(user_id):
    with closing(sqlite3.connect(DB_PATH)) as c:
        row = c.execute("SELECT lang FROM users WHERE user_id = ?", (user_id,)).fetchone()
    return row[0] if row else None


def set_user_lang(user_id, lang, name="User"):
    with closing(sqlite3.connect(DB_PATH)) as c:
        c.execute("INSERT OR REPLACE INTO users (user_id, full_name, lang) VALUES (?, ?, ?)", (user_id, name, lang))
        c.commit()


def lang_of(user_id):
    return get_user_lang(user_id) or "uz_lat"


def T(user_id, key):
    return tr_lang(lang_of(user_id), key)


def lang_keyboard():
    b = InlineKeyboardBuilder()
    b.button(text="🇺🇿 O'zbekcha", callback_data="lang:uz_lat")
    b.button(text="🇺🇿 Ўзбекча", callback_data="lang:uz_cyr")
    b.button(text="🇷🇺 Русский", callback_data="lang:ru")
    b.button(text="🇬🇧 English", callback_data="lang:en")
    b.adjust(2)
    return b.as_markup()


def reply_kb(user_id, rows):
    b = ReplyKeyboardBuilder()
    for row in rows:
        for key in row:
            b.button(text=T(user_id, key))
    b.adjust(*[len(r) for r in rows])
    return b.as_markup(resize_keyboard=True)


# ---------------- Shriftlar ----------------
FONT_SPECS = {
    "sans": ("DejaVuSans.ttf", [
        "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf",
        "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"]),
    "sans_bold": ("DejaVuSans-Bold.ttf", [
        "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSans-Bold.ttf",
        "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans-Bold.ttf"]),
    "serif": ("DejaVuSerif.ttf", [
        "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSerif.ttf"]),
    "serif_bold": ("DejaVuSerif-Bold.ttf", [
        "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSerif-Bold.ttf"]),
    "oblique": ("DejaVuSans-Oblique.ttf", [
        "https://raw.githubusercontent.com/matplotlib/matplotlib/main/lib/matplotlib/mpl-data/fonts/ttf/DejaVuSans-Oblique.ttf"]),
}
HAND_CANDIDATES = [
    ("handwriting.ttf", []),  # o'zingiz qo'yishingiz mumkin: fonts/handwriting.ttf
    ("Caveat.ttf", ["https://github.com/google/fonts/raw/main/ofl/caveat/Caveat%5Bwght%5D.ttf"]),
    ("MarckScript-Regular.ttf", ["https://github.com/google/fonts/raw/main/ofl/marckscript/MarckScript-Regular.ttf"]),
    ("BadScript-Regular.ttf", ["https://github.com/google/fonts/raw/main/ofl/badscript/BadScript-Regular.ttf"]),
]
SYSTEM_FONT_DIRS = [
    FONT_DIR, "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/dejavu", "/usr/share/fonts/TTF",
    "/usr/share/fonts/truetype", "/usr/local/share/fonts",
]
_font_lock = threading.Lock()
_font_cache = {}
_font_last_try = [0.0]


def _is_font(data: bytes) -> bool:
    return len(data) > 20000 and data[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO", b"ttcf")


def locate_font(filename):
    for d in SYSTEM_FONT_DIRS:
        p = os.path.join(d, filename)
        if os.path.isfile(p):
            return p
    try:
        import matplotlib
        p = os.path.join(os.path.dirname(matplotlib.__file__), "mpl-data", "fonts", "ttf", filename)
        if os.path.isfile(p):
            return p
    except Exception:
        pass
    for root in ("/usr/share/fonts", "/usr/local/share/fonts"):
        for dirpath, _, files in os.walk(root):
            if filename in files:
                return os.path.join(dirpath, filename)
    return None


def _download_font(filename, urls):
    for url in urls:
        try:
            data = urllib.request.urlopen(url, timeout=25).read()
            if _is_font(data):
                path = os.path.join(FONT_DIR, filename)
                with open(path, "wb") as f:
                    f.write(data)
                return path
        except Exception as e:
            logging.info("Shrift yuklanmadi (%s): %s", url, e)
    return None


def ensure_fonts():
    """Shriftlarni topadi yoki internetdan yuklaydi. Natija: dict (topilmaganlari None)."""
    with _font_lock:
        if _font_cache.get("sans"):
            return _font_cache
        if _font_cache and time.time() - _font_last_try[0] < 300:
            return _font_cache  # yaqinda urinib ko'rilgan, qayta sekinlashtirmaymiz
        _font_last_try[0] = time.time()
        res = {}
        for key, (fname, urls) in FONT_SPECS.items():
            res[key] = locate_font(fname) or _download_font(fname, urls)
        hand = None
        for fname, urls in HAND_CANDIDATES:
            hand = locate_font(fname) or (_download_font(fname, urls) if urls else None)
            if hand:
                break
        res["hand"] = hand or res.get("oblique") or res.get("sans")
        _font_cache.clear()
        _font_cache.update(res)
        logging.info("Shriftlar: %s", {k: bool(v) for k, v in res.items()})
        return _font_cache


_rl_fonts = {}


def rl_fonts(kind="sans"):
    """reportlab uchun (oddiy, qalin) shrift nomlari."""
    if kind in _rl_fonts:
        return _rl_fonts[kind]
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    paths = ensure_fonts()
    reg, bold = paths.get(kind), paths.get(kind + "_bold")
    names = ("Helvetica", "Helvetica-Bold")
    if reg and bold:
        try:
            n1, n2 = f"RL-{kind}", f"RL-{kind}-B"
            pdfmetrics.registerFont(TTFont(n1, reg))
            pdfmetrics.registerFont(TTFont(n2, bold))
            pdfmetrics.registerFontFamily(n1, normal=n1, bold=n2, italic=n1, boldItalic=n2)
            names = (n1, n2)
        except Exception as e:
            logging.warning("reportlab shrift xatosi: %s", e)
    else:
        logging.warning("Unicode shrift topilmadi — kirill/maxsus harflar to'g'ri chiqmasligi mumkin")
    _rl_fonts[kind] = names
    return names


# ---------------- Yordamchilar ----------------
def sess(uid):
    return user_sessions.setdefault(uid, {"mode": None, "files": []})


def reset(uid):
    s = user_sessions.get(uid)
    if s:
        for f in s.get("files", []):
            safe_remove(f)
    user_sessions[uid] = {"mode": None, "files": []}


def tmp_path(uid, ext):
    return os.path.join(TMP_DIR, f"{uid}_{uuid.uuid4().hex[:8]}{ext}")


def safe_remove(*paths):
    for p in paths:
        try:
            if p and os.path.exists(p):
                os.remove(p)
        except OSError:
            pass


async def run_blocking(fn, *args):
    async with HEAVY:
        return await asyncio.to_thread(fn, *args)


async def download_to(file_id, path):
    await bot.download(file_id, destination=path)


def esc(s):
    return html.escape(str(s), quote=False)


def parse_count(text, words):
    m = re.search(r"(\d{1,3})\s*[- ]?\s*(?:ta\s+)?(?:%s)" % "|".join(words), text, flags=re.I)
    return int(m.group(1)) if m else None


# =====================================================================
#                       FAYL O'ZGARTIRISH FUNKSIYALARI
# =====================================================================
def images_to_pdf(paths, out):
    imgs = []
    for p in paths:
        im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
        if max(im.size) > 3000:
            im.thumbnail((3000, 3000))
        imgs.append(im)
    if not imgs:
        raise ValueError("rasm yo'q")
    imgs[0].save(out, save_all=True, append_images=imgs[1:], resolution=150.0)


def merge_pdfs(paths, out):
    from pypdf import PdfWriter
    w = PdfWriter()
    for p in paths:
        w.append(p)
    with open(out, "wb") as f:
        w.write(f)
    w.close()


def pdf_to_images(pdf_path, out_dir, max_pages=30):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        import pymupdf as fitz
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    total = doc.page_count
    paths = []
    for i in range(min(total, max_pages)):
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
        p = os.path.join(out_dir, f"page_{i + 1:03d}.png")
        pix.save(p)
        paths.append(p)
    doc.close()
    return paths, total


def pdf_to_docx(pdf_path, out):
    from pypdf import PdfReader
    from docx import Document
    reader = PdfReader(pdf_path)
    d = Document()
    has_text = False
    for i, page in enumerate(reader.pages):
        text = (page.extract_text() or "").strip()
        if text:
            has_text = True
            d.add_paragraph(text)
        if i < len(reader.pages) - 1:
            d.add_page_break()
    if not has_text:
        raise ValueError("PDF ichida matn topilmadi (skaner qilingan bo'lishi mumkin)")
    d.save(out)


def _rl_inline(s):
    return html.escape(s, quote=False)


def docx_to_pdf(path, out):
    from docx import Document
    from docx.table import Table as DTable
    from docx.text.paragraph import Paragraph as DPar
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    reg, bold = rl_fonts("sans")
    base = ParagraphStyle("b", fontName=reg, fontSize=11, leading=15, spaceAfter=4)
    heading = ParagraphStyle("h", parent=base, fontName=bold, fontSize=15, leading=19, spaceBefore=8, spaceAfter=6)
    aligns = {1: TA_CENTER, 2: TA_RIGHT, 3: TA_JUSTIFY}

    def runs_markup(par):
        parts = []
        for r in par.runs:
            t = _rl_inline(r.text).replace("\n", "<br/>")
            if not t:
                continue
            if r.bold:
                t = f"<b>{t}</b>"
            if r.italic:
                t = f"<i>{t}</i>"
            if r.underline:
                t = f"<u>{t}</u>"
            parts.append(t)
        return "".join(parts) or _rl_inline(par.text)

    doc = Document(path)
    story = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            par = DPar(child, doc)
            text = par.text.strip()
            if not text:
                story.append(Spacer(1, 6))
                continue
            sname = (par.style.name or "") if par.style is not None else ""
            if sname.startswith("Heading") or sname == "Title":
                story.append(Paragraph(_rl_inline(text), heading))
            else:
                al = aligns.get(int(par.alignment) if par.alignment is not None else 0, TA_LEFT)
                st = ParagraphStyle("p", parent=base, alignment=al)
                story.append(Paragraph(runs_markup(par), st))
        elif child.tag.endswith("}tbl"):
            tbl = DTable(child, doc)
            cell_style = ParagraphStyle("c", parent=base, fontSize=9, leading=11)
            data = [[Paragraph(_rl_inline(c.text), cell_style) for c in row.cells] for row in tbl.rows]
            if data and data[0]:
                t = Table(data, repeatRows=1)
                t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
                story.append(t)
                story.append(Spacer(1, 8))
    if not story:
        story.append(Paragraph("(bo'sh hujjat)", base))
    SimpleDocTemplate(out, pagesize=A4, leftMargin=2.2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm).build(story)


def xlsx_to_pdf(path, out):
    import openpyxl
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

    reg, bold = rl_fonts("sans")
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows(values_only=True):
            cells = ["" if c is None else str(c)[:300] for c in row[:20]]
            if any(c.strip() for c in cells):
                rows.append(cells)
            if len(rows) >= 3000:
                break
        if rows:
            sheets.append((ws.title, rows))
    if not sheets:
        raise ValueError("Excel faylda ma'lumot topilmadi")
    max_cols = max(len(r) for _, rows in sheets for r in rows)
    page = landscape(A4) if max_cols > 6 else A4
    avail = page[0] - 3 * cm
    fsize = 9 if max_cols <= 6 else (8 if max_cols <= 10 else 6.5)
    head = ParagraphStyle("h", fontName=bold, fontSize=13, leading=16, spaceAfter=6)
    cell = ParagraphStyle("c", fontName=reg, fontSize=fsize, leading=fsize + 2)
    story = []
    for si, (title, rows) in enumerate(sheets):
        if si:
            story.append(PageBreak())
        story.append(Paragraph(_rl_inline(title), head))
        n = max(len(r) for r in rows)
        data = [[Paragraph(_rl_inline(c), cell) for c in (r + [""] * (n - len(r)))] for r in rows]
        t = Table(data, colWidths=[avail / n] * n, repeatRows=1)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF2F7")),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))
    SimpleDocTemplate(out, pagesize=page, leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm).build(story)


def enhance_image(path, out):
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    if max(im.size) < 1600:
        k = 2 if max(im.size) < 900 else 1.5
        im = im.resize((int(im.width * k), int(im.height * k)), Image.LANCZOS)
    im = ImageOps.autocontrast(im, cutoff=1)
    im = ImageEnhance.Contrast(im).enhance(1.15)
    im = ImageEnhance.Sharpness(im).enhance(1.6)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.2, percent=90, threshold=3))
    im.save(out, "JPEG", quality=95)
    return im.size


# =====================================================================
#                           QO'LDA YOZISH
# =====================================================================
def render_handwriting(text, max_pages=12, seed=None):
    rnd = random.Random(seed)
    fonts = ensure_fonts()
    fpath = fonts.get("hand")
    W, H = 1240, 1754
    size, line_h, top = 50, 66, 175
    left, right = 175, 85
    if fpath and "caveat" in os.path.basename(fpath).lower():
        size = 60  # Caveat kichikroq ko'rinadi
    try:
        font = ImageFont.truetype(fpath, size) if fpath else ImageFont.load_default(size)
    except Exception:
        font = ImageFont.load_default(size)
    ink = (22, 33, 112)
    text = re.sub(r"[ʻʼ‘’`]", "'", text.replace("\t", "    "))

    def new_page():
        img = Image.new("RGB", (W, H), (253, 251, 244))
        d = ImageDraw.Draw(img)
        y = top
        while y < H - 90:
            d.line([(55, y), (W - 55, y)], fill=(176, 198, 228), width=2)
            y += line_h
        d.line([(140, 0), (140, H)], fill=(238, 130, 130), width=3)
        d.line([(148, 0), (148, H)], fill=(238, 130, 130), width=1)
        return img

    pages, img = [], new_page()
    y, x = top, left
    wrote = [False]
    space = font.getlength(" ")
    base = int(size * 1.35)

    def put_word(w, wx, wy):
        wl = int(font.getlength(w)) + 40
        tmp = Image.new("RGBA", (wl, size * 2), (0, 0, 0, 0))
        ImageDraw.Draw(tmp).text((20, base), w, font=font, fill=ink + (255,), anchor="ls")
        tmp = tmp.rotate(rnd.uniform(-1.8, 1.8), resample=Image.BICUBIC)
        img.paste(tmp, (int(wx) - 20, int(wy) - base - 6 + rnd.randint(-3, 3)), tmp)
        wrote[0] = True

    def next_line():
        nonlocal img, y, x
        y += line_h
        x = left
        if y > H - 110:
            pages.append(img)
            img = new_page()
            wrote[0] = False
            y = top

    paragraphs = text.split("\n")
    for para in paragraphs:
        words = para.split()
        if not words:
            next_line()
            continue
        for w in words:
            while font.getlength(w) > W - left - right:  # juda uzun so'z
                cut = max(1, int(len(w) * (W - left - right) / font.getlength(w)) - 1)
                if x > left:
                    next_line()
                put_word(w[:cut], x, y)
                w = w[cut:]
                next_line()
            wl = font.getlength(w)
            if x + wl > W - right and x > left:
                next_line()
            put_word(w, x, y)
            x += wl + space * rnd.uniform(0.9, 1.35)
        next_line()
        if len(pages) >= max_pages:
            break
    if wrote[0] or not pages:
        pages.append(img)
    return pages[:max_pages]


# =====================================================================
#                        AI (GROQ) VA HUJJATLAR
# =====================================================================
class AIError(Exception):
    pass


GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODELS = []
for _m in (os.getenv("GROQ_MODEL", "").strip(), "openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.6-27b"):
    if _m and _m not in GROQ_MODELS:
        GROQ_MODELS.append(_m)


async def groq_chat(messages, max_tokens=6000, temperature=0.6):
    if not KEY_GROQ:
        raise AIError("GROQ_API_KEY topilmadi")
    headers = {"Authorization": f"Bearer {KEY_GROQ}", "Content-Type": "application/json"}
    last = "javob yo'q"
    timeout = aiohttp.ClientTimeout(total=150)
    async with aiohttp.ClientSession(timeout=timeout) as s:
        for model in GROQ_MODELS:
            payload = {"model": model, "messages": messages, "temperature": temperature, "max_completion_tokens": max_tokens}
            if model.startswith("openai/gpt-oss"):
                payload["reasoning_effort"] = "low"
            for _ in range(3):
                try:
                    async with s.post(GROQ_URL, json=payload, headers=headers) as r:
                        status, text = r.status, await r.text()
                except (asyncio.TimeoutError, aiohttp.ClientError) as e:
                    last = f"{model}: {type(e).__name__}"
                    await asyncio.sleep(1)
                    continue
                if status == 200:
                    try:
                        content = (json.loads(text).get("choices") or [{}])[0].get("message", {}).get("content") or ""
                    except ValueError:
                        last = f"{model}: noto'g'ri javob"
                        break
                    content = re.sub(r"<think>.*?</think>", "", content, flags=re.S).strip()
                    if content:
                        return content
                    last = f"{model}: bo'sh javob"
                    break
                last = f"{model} [{status}] {text[:120]}"
                if status == 401:
                    raise AIError("Groq kaliti noto'g'ri (401)")
                if status == 429:
                    m = re.search(r"try again in ([\d.]+)s", text)
                    await asyncio.sleep(min(float(m.group(1)) + 0.5 if m else 5, 25))
                    continue
                if status >= 500:
                    await asyncio.sleep(2)
                    continue
                break
    raise AIError(last)


def extract_json(text):
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    a, b = text.find("{"), text.rfind("}")
    if a < 0 or b <= a:
        raise ValueError("JSON topilmadi")
    return json.loads(text[a:b + 1])


async def groq_json(messages, **kw):
    last = None
    for _ in range(2):
        txt = await groq_chat(messages, **kw)
        try:
            return extract_json(txt)
        except ValueError as e:
            last = e
            messages = messages + [{"role": "assistant", "content": txt}, {"role": "user", "content": "Return ONLY valid JSON, nothing else."}]
    raise AIError(f"AI javobi noto'g'ri formatda: {last}")


# --- Taqdimot ---
async def gen_slides(topic, lang, n):
    sys_p = (
        f"You are a presentation designer. Write in {LANG_NAMES[lang]}. Return ONLY valid JSON (no markdown fences) with this schema: "
        '{"title": str, "subtitle": str, "slides": [{"title": str, "bullets": [str], "notes": str}]}. '
        f"Create exactly {n} content slides (not counting the title slide). The first content slide is an agenda/outline, the last is a conclusion. "
        "Each slide has 3-5 concise bullets (max 14 words each), plain text only (no markdown). notes = 1-2 sentences for the speaker."
    )
    data = await groq_json([{"role": "system", "content": sys_p}, {"role": "user", "content": topic}], max_tokens=7000, temperature=0.5)
    slides = []
    for s in data.get("slides") or []:
        if isinstance(s, dict) and s.get("title"):
            bullets = [str(b).strip() for b in (s.get("bullets") or []) if str(b).strip()][:6]
            slides.append({"title": str(s["title"]).strip(), "bullets": bullets, "notes": str(s.get("notes") or "").strip()})
    if not slides:
        raise AIError("AI slaydlarni yarata olmadi")
    return {"title": str(data.get("title") or topic)[:150], "subtitle": str(data.get("subtitle") or "")[:200], "slides": slides[:25]}


PALETTES = [
    {"bg": (17, 38, 76), "accent": (255, 183, 3), "light": (245, 248, 252), "text": (30, 41, 59)},
    {"bg": (20, 83, 70), "accent": (255, 209, 102), "light": (244, 250, 247), "text": (28, 45, 40)},
    {"bg": (88, 28, 135), "accent": (251, 191, 36), "light": (250, 245, 255), "text": (46, 16, 71)},
    {"bg": (127, 29, 29), "accent": (253, 224, 71), "light": (254, 247, 245), "text": (60, 20, 20)},
    {"bg": (30, 41, 59), "accent": (56, 189, 248), "light": (241, 245, 249), "text": (15, 23, 42)},
]


def build_pptx(data, out, author=None):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE

    pal = random.choice(PALETTES)
    rgb = lambda t: RGBColor(*t)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    def rect(slide, x, y, w, h, color):
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        return shp

    def textbox(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size, r.font.bold = Pt(size), bold
        r.font.color.rgb = rgb(color)
        return tf

    # sarlavha slaydi
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, H, pal["bg"])
    rect(s, Inches(0.9), Inches(3.55), Inches(1.6), Inches(0.12), pal["accent"])
    textbox(s, Inches(0.9), Inches(1.5), Inches(11.5), Inches(2.0), data["title"], 44, (255, 255, 255), True, anchor=MSO_ANCHOR.BOTTOM)
    sub = data.get("subtitle") or ""
    if author:
        sub = (sub + "\n" if sub else "") + str(author)
    if sub:
        textbox(s, Inches(0.9), Inches(3.9), Inches(11.5), Inches(1.8), sub, 22, (220, 228, 240))

    total = len(data["slides"])
    for i, sl in enumerate(data["slides"], 1):
        s = prs.slides.add_slide(blank)
        rect(s, 0, 0, W, H, pal["light"])
        rect(s, 0, 0, W, Inches(1.35), pal["bg"])
        rect(s, 0, Inches(1.35), W, Inches(0.07), pal["accent"])
        textbox(s, Inches(0.7), Inches(0.12), Inches(11.9), Inches(1.1), sl["title"], 30, (255, 255, 255), True, anchor=MSO_ANCHOR.MIDDLE)
        bullets = sl["bullets"] or [""]
        n_chars = sum(len(b) for b in bullets)
        size = 26 if n_chars < 280 else 22 if n_chars < 420 else 19
        tb = s.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(11.6), Inches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True
        for j, b in enumerate(bullets):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.space_after = Pt(10)
            r = p.add_run()
            r.text = "•  " + b
            r.font.size = Pt(size)
            r.font.color.rgb = rgb(pal["text"])
        textbox(s, Inches(12.0), Inches(6.95), Inches(1.0), Inches(0.4), f"{i}/{total}", 12, (120, 130, 145), align=PP_ALIGN.RIGHT)
        if sl.get("notes"):
            s.notes_slide.notes_text_frame.text = sl["notes"]
    prs.save(out)


# --- Matnli hujjatlar ---
KINDS = {
    "ai_referat": {"name": "Referat / Mustaqil ish", "pages": 8, "cover": True, "justify": True,
                   "structure": "Title; Contents (short plan); Introduction; 3 main sections each with 2-3 subsections; Conclusion; List of references (6-8 items)"},
    "ai_kurs": {"name": "Kurs ishi", "pages": 15, "cover": True, "justify": True,
                "structure": "Title; Contents; Introduction (relevance, aim, tasks, object, subject, methods); Chapter I (theory, 3 paragraphs); Chapter II (analysis/practice, 3 paragraphs); Chapter III (recommendations); Conclusion; List of references (8-10 items)"},
    "ai_research": {"name": "Tadqiqot (ilmiy ish)", "pages": 8, "cover": True, "justify": True,
                    "structure": "Title; Abstract; Keywords; Introduction; Literature review; Methodology; Results and discussion; Conclusion; References"},
    "ai_dars": {"name": "Dars ishlanma", "pages": 3, "cover": False, "justify": False,
                "structure": "Lesson plan: subject, grade, topic, lesson type, aims (educational, developmental, upbringing), equipment, lesson stages with minutes (organisation, review, new material, practice, consolidation, assessment, homework), questions and tasks"},
    "ai_resume": {"name": "Rezyume / Obektivka", "pages": 1, "cover": False, "justify": False,
                  "structure": "A clean one-page resume (summary, education, experience, skills, languages, contacts). If the user asks for an 'obektivka', write the standard Uzbek 'MA'LUMOTNOMA' personal data sheet (tug'ilgan yili, joyi, millati, ma'lumoti, mutaxassisligi, chet tillari, mehnat faoliyati)."},
    "ai_other": {"name": "Hujjat", "pages": 3, "cover": False, "justify": True,
                 "structure": "Choose a suitable structure for the requested document type (essay, article, thesis, application, letter, test questions, script...)"},
}


def strip_md(s):
    return re.sub(r"\*\*|__|`", "", s).strip()


def parse_markdown(md):
    title, blocks = None, []
    for raw in md.splitlines():
        line = raw.rstrip()
        if not line.strip() or line.strip().startswith("```"):
            continue
        m = re.match(r"^(#{1,4})\s+(.*)$", line)
        if m:
            text = strip_md(m.group(2))
            if len(m.group(1)) == 1 and title is None:
                title = text
            else:
                blocks.append(("h1" if len(m.group(1)) <= 2 else "h2", text))
            continue
        m = re.match(r"^\s*[-*•]\s+(.*)$", line)
        if m:
            blocks.append(("li", m.group(1).strip()))
            continue
        if re.match(r"^\s*[-=_]{3,}\s*$", line):
            continue
        blocks.append(("p", line.strip()))
    return title or "Hujjat", blocks


async def gen_text_doc(kind, topic, lang, pages):
    cfg = KINDS[kind]
    words = max(150, min(pages * 300, 3600)) if kind != "ai_resume" else 350
    system = (
        f"You are an expert academic writer. Write in {LANG_NAMES[lang]}; translate all headings into that language. "
        "Use Markdown: '# ' for the document title (once), '## ' for headings, '### ' for subheadings, '- ' for bullet points, **bold** sparingly. "
        "Write real, specific, well-organised content (no placeholders, no meta commentary, no code fences). "
        "Do not invent fake statistics; references must be plausible well-known works."
    )
    if kind in ("ai_referat", "ai_kurs", "ai_research") and words > 1800:
        plan = await groq_json([
            {"role": "system", "content": system},
            {"role": "user", "content": (
                f"Document type: {cfg['name']}. Topic: {topic}. Structure: {cfg['structure']}.\n"
                'Return ONLY JSON: {"title": str, "sections": [str]} with 6-8 section headings in order '
                "(introduction first, conclusion and references last).")},
        ], max_tokens=1500, temperature=0.4)
        sections = [str(x) for x in (plan.get("sections") or []) if str(x).strip()][:8]
        title = str(plan.get("title") or topic)
        if len(sections) < 3:
            raise AIError("AI reja tuza olmadi")
        per = max(250, words // len(sections))
        parts = [f"# {title}"]
        for i, sec in enumerate(sections, 1):
            body = await groq_chat([
                {"role": "system", "content": system},
                {"role": "user", "content": (
                    f"Document: {cfg['name']} on «{topic}». Full outline: {' | '.join(sections)}.\n"
                    f"Write ONLY section {i}/{len(sections)}: «{sec}» (about {per} words). "
                    f"Start with the line '## {sec}'. Do not repeat other sections.")},
            ], max_tokens=3500)
            parts.append(body.strip())
        return "\n\n".join(parts)
    user = (
        f"Document type: {cfg['name']}.\nTopic / user request: {topic}\n"
        f"Structure: {cfg['structure']}.\nTarget length: about {words} words."
    )
    return await groq_chat([{"role": "system", "content": system}, {"role": "user", "content": user}], max_tokens=7000)


def _add_runs(par, text, size=None):
    from docx.shared import Pt
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        bold = part.startswith("**") and part.endswith("**") and len(part) > 4
        r = par.add_run(part[2:-2] if bold else part)
        r.bold = bold or None
        if size:
            r.font.size = Pt(size)


def build_docx(title, blocks, path, author=None, cover=True, justify=True):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    d = Document()
    for sec in d.sections:
        sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(14 if cover else 12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.paragraph_format.line_spacing = 1.5 if cover else 1.15
    base = 14 if cover else 12

    if cover:
        for _ in range(7):
            d.add_paragraph()
        p = d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(22)
        for _ in range(8):
            d.add_paragraph()
        if author:
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(author))
        p = d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(datetime.now().year))
        d.add_page_break()
    else:
        p = d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(base + 4)

    for typ, text in blocks:
        if typ == "h1":
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if cover else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(base + 2)
        elif typ == "h2":
            p = d.add_paragraph()
            p.paragraph_format.keep_with_next = True
            r = p.add_run(text)
            r.bold = True
        elif typ == "li":
            p = d.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            _add_runs(p, "•  " + text)
        else:
            p = d.add_paragraph()
            if justify:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(1.25)
            _add_runs(p, text)
    d.save(path)


def build_pdf_doc(title, blocks, path, author=None, cover=True, justify=True):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak

    reg, bold = rl_fonts("serif")
    fs = 12 if cover else 10.5
    body = ParagraphStyle("body", fontName=reg, fontSize=fs, leading=fs * 1.5, alignment=TA_JUSTIFY if justify else TA_LEFT,
                          firstLineIndent=(0.9 * cm if justify else 0), spaceAfter=5)
    h1 = ParagraphStyle("h1", fontName=bold, fontSize=fs + 3, leading=fs + 7, alignment=TA_CENTER if cover else TA_LEFT, spaceBefore=10, spaceAfter=6, keepWithNext=1)
    h2 = ParagraphStyle("h2", fontName=bold, fontSize=fs + 1, leading=fs + 5, spaceBefore=6, spaceAfter=4, keepWithNext=1)
    li = ParagraphStyle("li", parent=body, firstLineIndent=0, leftIndent=0.9 * cm, bulletIndent=0.3 * cm, alignment=TA_LEFT)

    def mk(s):
        return re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", html.escape(s, quote=False))

    story = []
    if cover:
        story += [Spacer(1, 6 * cm),
                  Paragraph(f"<b>{html.escape(title, quote=False)}</b>", ParagraphStyle("t", fontName=bold, fontSize=22, leading=28, alignment=TA_CENTER)),
                  Spacer(1, 8 * cm)]
        if author:
            story.append(Paragraph(html.escape(str(author), quote=False), ParagraphStyle("a", fontName=reg, fontSize=13, alignment=TA_RIGHT)))
        story += [Paragraph(str(datetime.now().year), ParagraphStyle("y", fontName=reg, fontSize=13, alignment=TA_CENTER)), PageBreak()]
    else:
        story.append(Paragraph(f"<b>{html.escape(title, quote=False)}</b>", ParagraphStyle("t", fontName=bold, fontSize=fs + 5, leading=fs + 9, alignment=TA_CENTER, spaceAfter=8)))
    for typ, text in blocks:
        if typ == "h1":
            story.append(Paragraph(mk(text), h1))
        elif typ == "h2":
            story.append(Paragraph(mk(text), h2))
        elif typ == "li":
            story.append(Paragraph(mk(text), li, bulletText="•"))
        else:
            story.append(Paragraph(mk(text), body))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont(reg, 9)
        canvas.drawCentredString(A4[0] / 2, 1.1 * cm, str(doc_.page))
        canvas.restoreState()

    SimpleDocTemplate(path, pagesize=A4, leftMargin=3 * cm, rightMargin=1.5 * cm, topMargin=2 * cm, bottomMargin=2 * cm, title=title).build(
        story, onFirstPage=(lambda c, d: None) if cover else footer, onLaterPages=footer)


# =====================================================================
#                             HANDLERLAR
# =====================================================================
ai_usage = {}


def check_ai_limit(uid):
    if uid in ADMIN_IDS:
        return True
    day = datetime.now().strftime("%Y-%m-%d")
    d, n = ai_usage.get(uid, (day, 0))
    if d != day:
        n = 0
    if n >= AI_DAILY_LIMIT:
        return False
    ai_usage[uid] = (day, n + 1)
    return True


async def show_main(message, uid):
    reset(uid)
    await message.answer(T(uid, "welcome"), reply_markup=reply_kb(uid, MAIN_ROWS))


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    uid = message.from_user.id
    reset(uid)
    if not get_user_lang(uid):
        await message.answer("🌐 Iltimos, tilni tanlang / Выберите язык:", reply_markup=lang_keyboard())
    else:
        await show_main(message, uid)


@dp.message(Command("lang"))
async def cmd_lang(message: types.Message):
    await message.answer(T(message.from_user.id, "choose_lang"), reply_markup=lang_keyboard())


@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message):
    await show_main(message, message.from_user.id)


@dp.callback_query(F.data.startswith("lang:"))
async def process_lang(callback: types.CallbackQuery):
    await callback.answer()
    lang = callback.data.split(":", 1)[1]
    if lang not in LANGS:
        return
    uid = callback.from_user.id
    set_user_lang(uid, lang, callback.from_user.full_name)
    reset(uid)
    try:
        await callback.message.edit_text("✅ OK")
    except Exception:
        pass
    await callback.message.answer(T(uid, "welcome"), reply_markup=reply_kb(uid, MAIN_ROWS))


@dp.message(F.text.in_(ALL_LABELS))
async def on_menu(message: types.Message):
    uid = message.from_user.id
    key = LABEL_TO_KEY[message.text]
    if key in ("cancel", "back"):
        if key == "back" and sess(uid)["mode"] is None:
            return await show_main(message, uid)
        return await show_main(message, uid)
    if key == "m_files":
        reset(uid)
        return await message.answer(T(uid, "files_menu"), reply_markup=reply_kb(uid, FILES_ROWS))
    if key == "m_ai":
        reset(uid)
        return await message.answer(T(uid, "ai_menu"), reply_markup=reply_kb(uid, AI_ROWS))
    if key == "help":
        return await message.answer(T(uid, "help_text"))
    if key == "ready":
        return await on_ready(message)
    if key in FILE_MODES or key in AI_MODES:
        reset(uid)
        sess(uid)["mode"] = key
        if key in AI_MODES and not KEY_GROQ:
            return await message.answer(T(uid, "no_ai"))
        rows = ACTION_ROWS if key in ("img_to_pdf", "pdf_merge") else CANCEL_ROWS
        await message.answer(T(uid, "p_" + key), reply_markup=reply_kb(uid, rows))


async def send_doc(message, path, caption=None, filename=None):
    await message.answer_document(types.FSInputFile(path, filename=filename) if filename else types.FSInputFile(path), caption=caption)


async def on_ready(message: types.Message):
    uid = message.from_user.id
    s = sess(uid)
    mode, files = s["mode"], list(s["files"])
    if not files or mode not in ("img_to_pdf", "pdf_merge"):
        return await message.answer(T(uid, "no_files"))
    status = await message.answer(T(uid, "working"))
    out = tmp_path(uid, ".pdf")
    try:
        if mode == "img_to_pdf":
            await run_blocking(images_to_pdf, files, out)
            await send_doc(message, out, "✅ PDF tayyor!", "rasmlar.pdf")
        else:
            await run_blocking(merge_pdfs, files, out)
            await send_doc(message, out, "✅ PDF fayllar birlashtirildi!", "birlashtirilgan.pdf")
        await status.delete()
    except Exception:
        logging.exception("ready xatosi")
        await status.edit_text(T(uid, "error"))
    finally:
        safe_remove(out, *files)
        s["files"] = []


def doc_kind(name, mime):
    n = (name or "").lower()
    if n.endswith(".pdf") or mime == "application/pdf":
        return "pdf"
    if n.endswith(".docx"):
        return "docx"
    if n.endswith(".xlsx") or n.endswith(".xlsm"):
        return "xlsx"
    if n.endswith(".txt"):
        return "txt"
    if (mime or "").startswith("image/") or n.endswith((".jpg", ".jpeg", ".png", ".webp", ".bmp")):
        return "image"
    return "other"


async def convert_pdf_to_word(message, uid, path):
    out = tmp_path(uid, ".docx")
    try:
        await run_blocking(pdf_to_docx, path, out)
        await send_doc(message, out, "✅ PDF → WORD tayyor!", "hujjat.docx")
    finally:
        safe_remove(out)


async def convert_pdf_to_images(message, uid, path):
    out_dir = os.path.join(TMP_DIR, f"{uid}_{uuid.uuid4().hex[:6]}")
    try:
        paths, total = await run_blocking(pdf_to_images, path, out_dir)
        for i in range(0, len(paths), 10):
            chunk = paths[i:i + 10]
            media = [types.InputMediaPhoto(media=types.FSInputFile(p)) for p in chunk]
            if len(media) == 1:
                await message.answer_photo(types.FSInputFile(chunk[0]))
            else:
                await message.answer_media_group(media=media)
        note = f"✅ {len(paths)} ta sahifa" + (f" (jami {total} tadan dastlabki {len(paths)} tasi)" if total > len(paths) else "")
        await message.answer(note)
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


async def convert_office(message, uid, path, kind):
    out = tmp_path(uid, ".pdf")
    try:
        await run_blocking(docx_to_pdf if kind == "docx" else xlsx_to_pdf, path, out)
        await send_doc(message, out, "✅ PDF tayyor!", "hujjat.pdf")
    finally:
        safe_remove(out)


async def do_enhance(message, uid, path):
    out = tmp_path(uid, ".jpg")
    try:
        size = await run_blocking(enhance_image, path, out)
        await send_doc(message, out, f"✨ Rasm tiniqlashtirildi! O'lchami: {size[0]}x{size[1]}", "tiniq.jpg")
    finally:
        safe_remove(out)


@dp.message(F.photo)
async def handle_photo(message: types.Message):
    uid = message.from_user.id
    s = sess(uid)
    mode = s["mode"]
    photo = message.photo[-1]
    if mode not in (None, "img_to_pdf", "edit_image"):
        return await message.answer(T(uid, "wrong_type").format(need="—"))
    path = tmp_path(uid, ".jpg")
    try:
        await download_to(photo.file_id, path)
        if mode == "edit_image":
            status = await message.answer(T(uid, "working"))
            try:
                await do_enhance(message, uid, path)
                await status.delete()
            finally:
                safe_remove(path)
            return
        s["mode"] = "img_to_pdf"
        s["files"].append(path)
        await message.answer(f"📥 Rasm qo'shildi ({len(s['files'])} ta). Yana yuboring yoki «{T(uid, 'ready')}» ni bosing.",
                             reply_markup=reply_kb(uid, ACTION_ROWS))
    except Exception:
        logging.exception("photo xatosi")
        safe_remove(path)
        await message.answer(T(uid, "error"))


@dp.message(F.document)
async def handle_document(message: types.Message):
    uid = message.from_user.id
    doc = message.document
    s = sess(uid)
    mode = s["mode"]
    kind = doc_kind(doc.file_name, doc.mime_type)
    if doc.file_size and doc.file_size > MAX_DOWNLOAD:
        return await message.answer(T(uid, "too_big"))

    need = {"img_to_pdf": "image", "edit_image": "image", "pdf_merge": "pdf", "pdf_to_img": "pdf", "pdf_to_word": "pdf",
            "word_to_pdf": "docx", "excel_to_pdf": "xlsx", "handwrite": "txt"}
    if mode in need and kind != need[mode]:
        return await message.answer(T(uid, "wrong_type").format(need=need[mode].upper()))
    if mode in AI_MODES:
        return await message.answer(T(uid, "wrong_type").format(need="text"))
    if mode is None and kind == "other":
        return await message.answer("⚠️ Qo'llab-quvvatlanadigan turlar: PDF, DOCX, XLSX, rasm, TXT.")

    ext = {"pdf": ".pdf", "docx": ".docx", "xlsx": ".xlsx", "txt": ".txt", "image": ".img", "other": ".bin"}[kind]
    path = tmp_path(uid, ext)
    keep = False
    try:
        await download_to(doc.file_id, path)
        eff = mode
        if eff is None:  # rejim tanlanmagan: fayl turiga qarab
            if kind == "docx":
                eff = "word_to_pdf"
            elif kind == "xlsx":
                eff = "excel_to_pdf"
            elif kind == "image":
                eff = "img_to_pdf"
            elif kind == "pdf":
                s["pending_pdf"] = path
                keep = True
                b = InlineKeyboardBuilder()
                b.button(text="📁 Word (.docx)", callback_data="pdfto:word")
                b.button(text="🖼 Rasm", callback_data="pdfto:img")
                b.adjust(2)
                return await message.answer(T(uid, "pick_pdf"), reply_markup=b.as_markup())
            elif kind == "txt":
                eff = "handwrite"

        if eff in ("img_to_pdf", "pdf_merge"):
            s["mode"] = eff
            s["files"].append(path)
            keep = True
            return await message.answer(f"📥 Qo'shildi ({len(s['files'])} ta). Yana yuboring yoki «{T(uid, 'ready')}» ni bosing.",
                                        reply_markup=reply_kb(uid, ACTION_ROWS))
        status = await message.answer(T(uid, "working"))
        try:
            if eff == "pdf_to_word":
                await convert_pdf_to_word(message, uid, path)
            elif eff == "pdf_to_img":
                await convert_pdf_to_images(message, uid, path)
            elif eff in ("word_to_pdf", "excel_to_pdf"):
                await convert_office(message, uid, path, "docx" if eff == "word_to_pdf" else "xlsx")
            elif eff == "edit_image":
                await do_enhance(message, uid, path)
            elif eff == "handwrite":
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    await do_handwrite(message, uid, f.read())
            await status.delete()
        except ValueError as e:
            await status.edit_text(f"❌ {esc(e)}")
        except Exception:
            logging.exception("convert xatosi")
            await status.edit_text(T(uid, "error"))
    except Exception:
        logging.exception("document xatosi")
        await message.answer(T(uid, "error"))
    finally:
        if not keep:
            safe_remove(path)


@dp.callback_query(F.data.startswith("pdfto:"))
async def pdf_choice(callback: types.CallbackQuery):
    await callback.answer()
    uid = callback.from_user.id
    s = sess(uid)
    path = s.pop("pending_pdf", None)
    if not path or not os.path.exists(path):
        return await callback.message.answer("⚠️ Faylni qaytadan yuboring.")
    status = await callback.message.answer(T(uid, "working"))
    try:
        if callback.data.endswith("word"):
            await convert_pdf_to_word(callback.message, uid, path)
        else:
            await convert_pdf_to_images(callback.message, uid, path)
        await status.delete()
    except ValueError as e:
        await status.edit_text(f"❌ {esc(e)}")
    except Exception:
        logging.exception("pdf choice xatosi")
        await status.edit_text(T(uid, "error"))
    finally:
        safe_remove(path)


async def do_handwrite(message, uid, text):
    text = (text or "").strip()
    if not text:
        return await message.answer(T(uid, "send_text"))
    text = text[:6000]
    pages = await run_blocking(render_handwriting, text)
    paths = []
    try:
        for i, im in enumerate(pages, 1):
            p = tmp_path(uid, f"_p{i}.jpg")
            im.save(p, "JPEG", quality=90)
            paths.append(p)
        pdf = tmp_path(uid, ".pdf")
        paths.append(pdf)
        pages[0].save(pdf, save_all=True, append_images=pages[1:], resolution=150.0)
        imgs = paths[:-1][:10]
        if len(imgs) == 1:
            await message.answer_photo(types.FSInputFile(imgs[0]))
        else:
            await message.answer_media_group(media=[types.InputMediaPhoto(media=types.FSInputFile(p)) for p in imgs])
        await send_doc(message, pdf, "✅ Qo'lyozma tayyor! (PDF)", "qolyozma.pdf")
    finally:
        safe_remove(*paths)


async def do_ai_doc(message, uid, mode, prompt):
    if not check_ai_limit(uid):
        return await message.answer(T(uid, "limit").format(n=AI_DAILY_LIMIT))
    lang = lang_of(uid)
    status = await message.answer(T(uid, "ai_wait"))
    files = []
    try:
        author = message.from_user.full_name
        if mode == "ai_slides":
            n = parse_count(prompt, ["slayd", "slide", "слайд", "слайдов"]) or 10
            n = max(4, min(n, 20))
            data = await gen_slides(prompt, lang, n)
            out = tmp_path(uid, ".pptx")
            files.append(out)
            await run_blocking(build_pptx, data, out, author)
            await send_doc(message, out, f"✅ Taqdimot tayyor! ({len(data['slides']) + 1} slayd)", "taqdimot.pptx")
        else:
            cfg = KINDS[mode]
            pages = parse_count(prompt, ["bet", "page", "pages", "стр", "страниц"]) or cfg["pages"]
            pages = max(1, min(pages, 25))
            md = await gen_text_doc(mode, prompt, lang, pages)
            title, blocks = parse_markdown(md)
            docx_p, pdf_p = tmp_path(uid, ".docx"), tmp_path(uid, ".pdf")
            files += [docx_p, pdf_p]
            await run_blocking(build_docx, title, blocks, docx_p, author, cfg["cover"], cfg["justify"])
            await run_blocking(build_pdf_doc, title, blocks, pdf_p, author, cfg["cover"], cfg["justify"])
            base = re.sub(r"[^\w\-]+", "_", title)[:40] or "hujjat"
            await send_doc(message, docx_p, f"✅ {cfg['name']} tayyor! (Word)", f"{base}.docx")
            await send_doc(message, pdf_p, "📄 PDF variant", f"{base}.pdf")
        await status.edit_text(T(uid, "ai_note"))
    except AIError as e:
        logging.warning("AI xatosi: %s", e)
        await status.edit_text(f"❌ AI xatosi: {esc(str(e)[:200])}\nBiroz kutib qayta urinib ko'ring.")
    except Exception:
        logging.exception("AI doc xatosi")
        await status.edit_text(T(uid, "error"))
    finally:
        safe_remove(*files)


@dp.message(F.text, ~F.text.startswith("/"))
async def handle_text(message: types.Message):
    uid = message.from_user.id
    mode = sess(uid)["mode"]
    text = message.text.strip()
    if mode in AI_MODES:
        return await do_ai_doc(message, uid, mode, text)
    if mode == "handwrite":
        status = await message.answer(T(uid, "working"))
        try:
            await do_handwrite(message, uid, text)
            await status.delete()
        except Exception:
            logging.exception("handwrite xatosi")
            await status.edit_text(T(uid, "error"))
        return
    await message.answer(T(uid, "menu_hint"))


async def main():
    asyncio.create_task(start_web_server())
    asyncio.create_task(asyncio.to_thread(ensure_fonts))
    while True:
        try:
            print("Bot polling boshlandi...")
            await bot.delete_webhook(drop_pending_updates=True)
            await dp.start_polling(bot, drop_pending_updates=True)
        except Exception as e:
            logging.error(f"Polling xatoligi: {e}")
            await asyncio.sleep(5)


if __name__ == "__main__":
    asyncio.run(main())
