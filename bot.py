import os
import sqlite3
import datetime
import logging
from aiogram import Bot, Dispatcher, F
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    ReplyKeyboardMarkup,
    KeyboardButton,
    LabeledPrice,
    PreCheckoutQuery,
    FSInputFile
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

from generator import generate_presentation_content, create_pptx_file

# Setări generale
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))
PAYMENT_PROVIDER_TOKEN = os.getenv("PAYMENT_PROVIDER_TOKEN", "")

# Canal și grup
CHANNEL_ID = os.getenv("CHANNEL_ID", "@kanalingiz_username") 
CHANNEL_URL = os.getenv("CHANNEL_URL", "https://t.me/kanalingiz_username")
GROUP_URL = os.getenv("GROUP_URL", "https://t.me/guruhingiz_username")

logging.basicConfig(level=logging.INFO)

# Baza de date SQLite
conn = sqlite3.connect("bot_database.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    full_name TEXT,
    created_at TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS promo_codes (
    code TEXT PRIMARY KEY,
    duration TEXT,
    used_count INTEGER DEFAULT 0
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS user_subscriptions (
    user_id INTEGER PRIMARY KEY,
    expire_at TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS user_promos (
    user_id INTEGER,
    code TEXT,
    used_at TEXT,
    PRIMARY KEY (user_id, code)
)
""")

cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('free_mode', 'true')")
cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('slide_price', '5000')")
conn.commit()

class PresentationState(StatesGroup):
    waiting_for_topic = State()
    waiting_for_script = State()

class PromoState(StatesGroup):
    waiting_for_code = State()

class AdminState(StatesGroup):
    waiting_for_promo_code = State()
    waiting_for_promo_duration = State()

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def is_free_mode():
    cursor.execute("SELECT value FROM settings WHERE key='free_mode'")
    res = cursor.fetchone()
    return res and res[0] == 'true'

def get_slide_price():
    cursor.execute("SELECT value FROM settings WHERE key='slide_price'")
    res = cursor.fetchone()
    return int(res[0]) if res else 5000

def has_active_subscription(user_id: int) -> bool:
    cursor.execute("SELECT expire_at FROM user_subscriptions WHERE user_id=?", (user_id,))
    res = cursor.fetchone()
    if not res:
        return False
    expire_str = res[0]
    if expire_str == "LIFETIME":
        return True
    try:
        expire_date = datetime.datetime.fromisoformat(expire_str)
        return datetime.datetime.now() < expire_date
    except Exception:
        return False

async def check_channel_sub(user_id: int) -> bool:
    if not CHANNEL_ID or CHANNEL_ID == "@kanalingiz_username":
        return True
    try:
        member = await bot.get_chat_member(chat_id=CHANNEL_ID, user_id=user_id)
        if member.status in ["creator", "administrator", "member"]:
            return True
        return False
    except Exception as e:
        logging.error(f"Eroare verificare canal: {e}")
        return True

def get_subscription_keyboard():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📢 Kanalimizga obuna bo'lish", url=CHANNEL_URL)],
        [InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_subscription")]
    ])
    return kb

def get_main_keyboard():
    kb = [
        [KeyboardButton(text="📊 Slayd yaratish")],
        [KeyboardButton(text="🎟 Promokod kiritish"), KeyboardButton(text="💬 Guruhimiz (Fikr bildirish)")],
        [KeyboardButton(text="ℹ️ Ma'lumot")]
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

@dp.message(Command("start"))
async def start_handler(message: Message):
    user_id = message.from_user.id
    full_name = message.from_user.full_name
    now_str = datetime.datetime.now().isoformat()

    cursor.execute("INSERT OR IGNORE INTO users (user_id, full_name, created_at) VALUES (?, ?, ?)", 
                   (user_id, full_name, now_str))
    conn.commit()

    is_subscribed = await check_channel_sub(user_id)
    if not is_subscribed:
        await message.answer(
            f"Salom, {full_name}!\n\n"
            f"⚠️ Botdan foydalanish uchun avval rasmiy **yangiliklar kanalimizga** obuna bo'ling. "
            f"Kanalda yangiliklar va tekin **promokodlar** berib boriladi!",
            reply_markup=get_subscription_keyboard(),
            parse_mode="Markdown"
        )
        return

    text = f"Salom, {full_name}! Prezentatsiya yaratuvchi botga xush kelibsiz.\n\nSlayd yaratish uchun quyidagi tugmani bosing:"
    await message.answer(text, reply_markup=get_main_keyboard())

@dp.callback_query(F.data == "check_subscription")
async def check_sub_callback(call: CallbackQuery):
    user_id = call.from_user.id
    is_subscribed = await check_channel_sub(user_id)

    if is_subscribed:
        await call.answer("✅ Rahmat! Obuna tasdiqlandi.", show_alert=True)
        await call.message.delete()
        await call.message.answer("Xush kelibsiz! Kerakli bo'limni tanlang:", reply_markup=get_main_keyboard())
    else:
        await call.answer("❌ Siz hali kanalga obuna bo'lmadingiz!", show_alert=True)

@dp.message(F.text == "💬 Guruhimiz (Fikr bildirish)")
async def group_info_handler(message: Message):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💬 Guruhga o'tish va fikr yozish", url=GROUP_URL)]
    ])
    await message.answer(
        "💬 **Bizning rasmiy muhokama guruhimiz!**\n\n"
        "Guruhda o'zingizning fikrlaringizni, takliflaringizni yozib qoldirishingiz va boshqa foydalanuvchilar bilan muloqot qilishingiz mumkin.",
        reply_markup=kb,
        parse_mode="Markdown"
    )

@dp.message(Command("admin"))
async def admin_handler(message: Message):
    if ADMIN_ID == 0 or message.from_user.id != ADMIN_ID:
        await message.answer("❌ Siz admin emassiz yoki ADMIN_ID sozlanmagan!")
        return

    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]

    free_status = "🟢 TEKIN" if is_free_mode() else "🔴 TO'LOVLI"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Rejim: {free_status}", callback_data="toggle_free_mode")],
        [InlineKeyboardButton(text="➕ Yangi Promokod yaratish", callback_data="add_promo")]
    ])

    await message.answer(
        f"⚙️ **Admin Paneli**\n\n"
        f"👥 Jami foydalanuvchilar: {total_users} ta\n"
        f"⚡️ Bot rejimi: {free_status}", 
        reply_markup=kb, 
        parse_mode="Markdown"
    )

@dp.callback_query(F.data == "toggle_free_mode")
async def toggle_free_mode_callback(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return

    current = is_free_mode()
    new_val = 'false' if current else 'true'
    cursor.execute("UPDATE settings SET value=? WHERE key='free_mode'", (new_val,))
    conn.commit()

    status_text = "🟢 TEKIN" if new_val == 'true' else "🔴 TO'LOVLI"
    await call.answer(f"Rejim o'zgartirildi: {status_text}")
    await admin_handler(call.message)

@dp.callback_query(F.data == "add_promo")
async def add_promo_start(call: CallbackQuery, state: FSMContext):
    if call.from_user.id != ADMIN_ID:
        return
    await state.set_state(AdminState.waiting_for_promo_code)
    await call.message.answer("Yangi promokod nomini (so'zini) kiriting:")

@dp.message(AdminState.waiting_for_promo_code)
async def process_promo_code_name(message: Message, state: FSMContext):
    code = message.text.strip().upper()
    await state.update_data(promo_code=code)
    await state.set_state(AdminState.waiting_for_promo_duration)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1 Kunlik", callback_data="dur_1d"), InlineKeyboardButton(text="1 Haftalik", callback_data="dur_1w")],
        [InlineKeyboardButton(text="1 Oylik", callback_data="dur_1m"), InlineKeyboardButton(text="Umrbod (Tekin)", callback_data="dur_life")]
    ])
    await message.answer(f"Promokod `{code}` uchun amal qilish muddatini tanlang:", reply_markup=kb, parse_mode="Markdown")

@dp.callback_query(AdminState.waiting_for_promo_duration, F.data.startswith("dur_"))
async def process_promo_duration(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    code = data.get("promo_code")
    dur_type = call.data.replace("dur_", "")

    cursor.execute("INSERT OR REPLACE INTO promo_codes (code, duration, used_count) VALUES (?, ?, 0)", (code, dur_type))
    conn.commit()

    await state.clear()
    await call.message.answer(f"✅ Promokod muvaffaqiyatli yaratildi!\n\nKod: `{code}`\nMuddat: {dur_type}", parse_mode="Markdown")

@dp.message(F.text == "🎟 Promokod kiritish")
async def enter_promo_start(message: Message, state: FSMContext):
    is_subscribed = await check_channel_sub(message.from_user.id)
    if not is_subscribed:
        await message.answer("⚠️ Botdan foydalanish uchun avval kanalga obuna bo'ling:", reply_markup=get_subscription_keyboard())
        return

    await state.set_state(PromoState.waiting_for_code)
    await message.answer("Sizda promokod bormi? Promokodni kiriting:")

@dp.message(PromoState.waiting_for_code)
async def process_promo_input(message: Message, state: FSMContext):
    code = message.text.strip().upper()
    user_id = message.from_user.id

    cursor.execute("SELECT duration FROM promo_codes WHERE code=?", (code,))
    res = cursor.fetchone()

    if not res:
        await message.answer("❌ Noto'g'ri yoki mavjud bo'lmagan promokod!")
        await state.clear()
        return

    cursor.execute("SELECT 1 FROM user_promos WHERE user_id=? AND code=?", (user_id, code))
    already_used = cursor.fetchone()

    if already_used:
        await message.answer("⚠️ Siz ushbu promokodni alaqachon ishlatgansiz!")
        await state.clear()
        return

    dur_type = res[0]
    now = datetime.datetime.now()

    if dur_type == "1d":
        expire = now + datetime.timedelta(days=1)
        expire_str = expire.isoformat()
    elif dur_type == "1w":
        expire = now + datetime.timedelta(weeks=1)
        expire_str = expire.isoformat()
    elif dur_type == "1m":
        expire = now + datetime.timedelta(days=30)
        expire_str = expire.isoformat()
    else:
        expire_str = "LIFETIME"

    cursor.execute("INSERT OR REPLACE INTO user_subscriptions (user_id
