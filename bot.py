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

# generator.py din proiectul tău
from generator import generate_presentation_content, create_pptx_file

# Setări generale
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))
PAYMENT_PROVIDER_TOKEN = os.getenv("PAYMENT_PROVIDER_TOKEN", "")

# Configurare Logging
logging.basicConfig(level=logging.INFO)

# Configurare Bază de Date SQLite
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

# Setări implicite
cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('free_mode', 'true')")
cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('slide_price', '5000')")
conn.commit()

# Stări FSM
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

# Funcții Ajutătoare
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
    expire_date = datetime.datetime.fromisoformat(expire_str)
    return datetime.datetime.now() < expire_date

# Meniu Principal
def get_main_keyboard():
    kb = [
        [KeyboardButton(text="📊 Slayd yaratish")],
        [KeyboardButton(text="🎟 Promokod kiritish"), KeyboardButton(text="ℹ️ Ma'lumot")]
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

# Start
@dp.message(Command("start"))
async def start_handler(message: Message):
    user_id = message.from_user.id
    full_name = message.from_user.full_name
    now_str = datetime.datetime.now().isoformat()

    cursor.execute("INSERT OR IGNORE INTO users (user_id, full_name, created_at) VALUES (?, ?, ?)", 
                   (user_id, full_name, now_str))
    conn.commit()

    text = f"Salom, {full_name}! Prezentatsiya yaratuvchi botga xush kelibsiz.\n\nSlayd yaratish uchun quyidagi tugmani bosing:"
    await message.answer(text, reply_markup=get_main_keyboard())

# Meniu Admin
@dp.message(Command("admin"))
async def admin_handler(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]

    free_status = "🟢 TEKIN" if is_free_mode() else "🔴 TO'LOVLI"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Rejimni o'zgartirish ({free_status})", callback_data="toggle_free_mode")],
        [InlineKeyboardButton(text="➕ Yangi Promokod yaratish", callback_data="add_promo")],
        [InlineKeyboardButton(text="📊 Statistika", callback_data="admin_stats")]
    ])

    await message.answer(f"⚙️ **Admin Paneli**\n\nJami foydalanuvchilar: {total_users}\nJoriy rejim: {free_status}", reply_markup=kb, parse_mode="Markdown")

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

# Utilizare Promokod
@dp.message(F.text == "🎟 Promokod kiritish")
async def enter_promo_start(message: Message, state: FSMContext):
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

    cursor.execute("INSERT OR REPLACE INTO user_subscriptions (user_id, expire_at) VALUES (?, ?)", (user_id, expire_str))
    cursor.execute("UPDATE promo_codes SET used_count = used_count + 1 WHERE code=?", (code,))
    conn.commit()

    await state.clear()
    await message.answer(f"🎉 Tabriklaymiz! Promokod faollashtirildi.\nSizga obuna taqdim etildi.")

# Creare Prezentare
@dp.message(F.text == "📊 Slayd yaratish")
async def start_presentation_flow(message: Message, state: FSMContext):
    await state.set_state(PresentationState.waiting_for_topic)
    await message.answer("Prezentatsiya mavzusini kiriting:")

@dp.message(PresentationState.waiting_for_topic)
async def process_topic(message: Message, state: FSMContext):
    topic = message.text.strip()
    await state.update_data(topic=topic)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Menda tayyor ssenariy bor", callback_data="has_script")],
        [InlineKeyboardButton(text="AI o'zi avtomatik yaratsin", callback_data="no_script")]
    ])
    await message.answer(f"Mavzu: **{topic}**\n\nSsenariyingiz bormi yoki AI o'zi yaratsinmi?", reply_markup=kb, parse_mode="Markdown")

@dp.callback_query(F.data.in_(["has_script", "no_script"]))
async def process_script_choice(call: CallbackQuery, state: FSMContext):
    if call.data == "has_script":
        await state.set_state(PresentationState.waiting_for_script)
        await call.message.answer("Ssenariyingizni matn ko'rinishida yuboring:")
    else:
        await generate_and_deliver(call.message, call.from_user.id, state)

@dp.message(PresentationState.waiting_for_script)
async def process_script_input(message: Message, state: FSMContext):
    await state.update_data(user_script=message.text.strip())
    await generate_and_deliver(message, message.from_user.id, state)

async def generate_and_deliver(message: Message, user_id: int, state: FSMContext):
    data = await state.get_data()
    topic = data.get("topic")
    user_script = data.get("user_script")

    # Verificare plată / acces gratuit
    if not is_free_mode() and not has_active_subscription(user_id):
        price = get_slide_price()
        if PAYMENT_PROVIDER_TOKEN:
            prices = [LabeledPrice(label="Prezentatsiya yaratish", amount=price * 100)]
            await bot.send_invoice(
                chat_id=user_id,
                title="Prezentatsiya uchun to'lov",
                description=f"'{topic}' mavzusida 6 ta slayd yaratish xizmati",
                provider_token=PAYMENT_PROVIDER_TOKEN,
                currency="UZS",
                prices=prices,
                start_parameter="create-presentation",
                payload=f"pres_{user_id}"
            )
            return
        else:
            await message.answer("⚠️ Hozirda to mezon tizimi faol emas. Iltimos adminga murojaat qiling.")
            await state.clear()
            return

    # Generare Prezentare
    msg = await message.answer("⏳ Prezentatsiya va rasmlar tayyorlanmoqda, iltimos kuting...")

    try:
        slides = await generate_presentation_content(topic, user_script)
        file_path = f"presentation_{user_id}.pptx"
        await create_pptx_file(slides, file_path)

        doc = FSInputFile(file_path)
        await message.answer_document(doc, caption=f"✅ Prezentatsiya tayyor!\nMavzu: {topic}")
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {e}")

    await state.clear()

@dp.pre_checkout_query()
async def process_pre_checkout(pre_checkout_query: PreCheckoutQuery):
    await bot.answer_pre_checkout_query(pre_checkout_query.id, ok=True)

@dp.message(F.successful_payment)
async def process_successful_payment(message: Message, state: FSMContext):
    await message.answer("✅ To'lov muvaffaqiyatli amalga oshirildi!")
    await generate_and_deliver(message, message.from_user.id, state)

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
