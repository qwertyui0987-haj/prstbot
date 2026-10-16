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
    FSInputFile
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

from generator import generate_presentation_content, create_pptx_file

# Sozlamalar
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# To'lov uchun karta ma'lumotlari
CARD_NUMBER = "8600 0000 0000 0000"  # O'zingizning karta raqamingizni kiriting
CARD_HOLDER = "FAZILOV F.F."          # Karta egasining ismi

# Canal va guruh
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
    max_uses INTEGER DEFAULT 0,
    expires_at TEXT,
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
conn.commit()

class PresentationState(StatesGroup):
    waiting_for_topic = State()
    waiting_for_script = State()

class PromoState(StatesGroup):
    waiting_for_code = State()

class PaymentState(StatesGroup):
    waiting_for_receipt = State()

class AdminState(StatesGroup):
    waiting_for_promo_code = State()
    waiting_for_promo_duration = State()
    waiting_for_promo_limit = State()
    waiting_for_promo_expiry = State()

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def is_free_mode():
    cursor.execute("SELECT value FROM settings WHERE key='free_mode'")
    res = cursor.fetchone()
    return res and res[0] == 'true'

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
        return member.status in ["creator", "administrator", "member"]
    except Exception as e:
        logging.error(f"Kanal obunasini tekshirishda xatolik: {e}")
        return True

def get_subscription_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📢 Kanalimizga obuna bo'lish", url=CHANNEL_URL)],
        [InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_subscription")]
    ])

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

    await message.answer(f"Salom, {full_name}! Prezentatsiya yaratuvchi botga xush kelibsiz.\n\nSlayd yaratish uchun quyidagi tugmani bosing:", reply_markup=get_main_keyboard())

@dp.callback_query(F.data == "check_subscription")
async def check_sub_callback(call: CallbackQuery):
    user_id = call.from_user.id
    if await check_channel_sub(user_id):
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

# ================= ADMIN PANEL =================
@dp.message(Command("admin"))
async def admin_handler(message: Message):
    if ADMIN_ID == 0 or message.from_user.id != ADMIN_ID:
        await message.answer("❌ Siz admin emassiz!")
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
    await call.message.answer("Yangi promokod nomini kiriting:")

@dp.message(AdminState.waiting_for_promo_code)
async def process_promo_code_name(message: Message, state: FSMContext):
    code = message.text.strip().upper()
    await state.update_data(promo_code=code)
    await state.set_state(AdminState.waiting_for_promo_duration)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1 Kunlik", callback_data="dur_1d"), InlineKeyboardButton(text="1 Haftalik", callback_data="dur_1w")],
        [InlineKeyboardButton(text="1 Oylik", callback_data="dur_1m"), InlineKeyboardButton(text="Umrbod", callback_data="dur_life")]
    ])
    await message.answer(f"Promokod `{code}` uchun obuna muddatini tanlang:", reply_markup=kb, parse_mode="Markdown")

@dp.callback_query(AdminState.waiting_for_promo_duration, F.data.startswith("dur_"))
async def process_promo_duration(call: CallbackQuery, state: FSMContext):
    dur_type = call.data.replace("dur_", "")
    await state.update_data(duration=dur_type)
    await state.set_state(AdminState.waiting_for_promo_limit)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="10 kishilik", callback_data="limit_10"), InlineKeyboardButton(text="50 kishilik", callback_data="limit_50")],
        [InlineKeyboardButton(text="100 kishilik", callback_data="limit_100"), InlineKeyboardButton(text="1000 kishilik", callback_data="limit_1000")],
        [InlineKeyboardButton(text="Cheksiz", callback_data="limit_0")]
    ])
    await call.message.answer("Promokod limitini tanlang:", reply_markup=kb)

@dp.callback_query(AdminState.waiting_for_promo_limit, F.data.startswith("limit_"))
async def process_promo_limit(call: CallbackQuery, state: FSMContext):
    limit_val = int(call.data.replace("limit_", ""))
    await state.update_data(max_uses=limit_val)
    await state.set_state(AdminState.waiting_for_promo_expiry)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1 Kun (24 soat)", callback_data="exp_1d"), InlineKeyboardButton(text="3 Kun", callback_data="exp_3d")],
        [InlineKeyboardButton(text="1 Hafta", callback_data="exp_1w"), InlineKeyboardButton(text="Muddatsiz", callback_data="exp_never")]
    ])
    await call.message.answer("Promokod tugash vaqtini tanlang:", reply_markup=kb)

@dp.callback_query(AdminState.waiting_for_promo_expiry, F.data.startswith("exp_"))
async def process_promo_expiry(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    code = data.get("promo_code")
    dur_type = data.get("duration")
    max_uses = data.get("max_uses")
    exp_type = call.data.replace("exp_", "")

    now = datetime.datetime.now()
    if exp_type == "1d":
        expires_at = (now + datetime.timedelta(days=1)).isoformat()
    elif exp_type == "3d":
        expires_at = (now + datetime.timedelta(days=3)).isoformat()
    elif exp_type == "1w":
        expires_at = (now + datetime.timedelta(weeks=1)).isoformat()
    else:
        expires_at = "NEVER"

    cursor.execute(
        "INSERT OR REPLACE INTO promo_codes (code, duration, max_uses, expires_at, used_count) VALUES (?, ?, ?, ?, 0)",
        (code, dur_type, max_uses, expires_at)
    )
    conn.commit()
    await state.clear()

    await call.message.answer(f"✅ Promokod yaratildi: `{code}`", parse_mode="Markdown")

# ================= PROMOKOD ISHLATISH =================
@dp.message(F.text == "🎟 Promokod kiritish")
async def enter_promo_start(message: Message, state: FSMContext):
    if not await check_channel_sub(message.from_user.id):
        await message.answer("⚠️ Avval kanalga obuna bo'ling:", reply_markup=get_subscription_keyboard())
        return

    await state.set_state(PromoState.waiting_for_code)
    await message.answer("Promokodni kiriting:")

@dp.message(PromoState.waiting_for_code)
async def process_promo_input(message: Message, state: FSMContext):
    code = message.text.strip().upper()
    user_id = message.from_user.id

    cursor.execute("SELECT duration, max_uses, expires_at, used_count FROM promo_codes WHERE code=?", (code,))
    res = cursor.fetchone()

    if not res:
        await message.answer("❌ Noto'g'ri promokod!")
        await state.clear()
        return

    dur_type, max_uses, expires_at, used_count = res
    now = datetime.datetime.now()

    if expires_at and expires_at != "NEVER":
        try:
            if now > datetime.datetime.fromisoformat(expires_at):
                await message.answer("❌ Ushbu promokodning muddati tugagan!")
                await state.clear()
                return
        except Exception:
            pass

    if max_uses and max_uses > 0 and used_count >= max_uses:
        await message.answer("❌ Ushbu promokod limiti tugagan!")
        await state.clear()
        return

    cursor.execute("SELECT 1 FROM user_promos WHERE user_id=? AND code=?", (user_id, code))
    if cursor.fetchone():
        await message.answer("⚠️ Siz ushbu promokodni ishlatgansiz!")
        await state.clear()
        return

    if dur_type == "1d":
        expire_str = (now + datetime.timedelta(days=1)).isoformat()
    elif dur_type == "1w":
        expire_str = (now + datetime.timedelta(weeks=1)).isoformat()
    elif dur_type == "1m":
        expire_str = (now + datetime.timedelta(days=30)).isoformat()
    else:
        expire_str = "LIFETIME"

    cursor.execute("INSERT OR REPLACE INTO user_subscriptions (user_id, expire_at) VALUES (?, ?)", (user_id, expire_str))
    cursor.execute("INSERT INTO user_promos (user_id, code, used_at) VALUES (?, ?, ?)", (user_id, code, now.isoformat()))
    cursor.execute("UPDATE promo_codes SET used_count = used_count + 1 WHERE code=?", (code,))
    conn.commit()

    await state.clear()
    await message.answer("🎉 Promokod faollashtirildi! Obuna taqdim etildi.")

# ================= PRESENTATION FLOW & PAYMENT =================
@dp.message(F.text == "📊 Slayd yaratish")
async def start_presentation_flow(message: Message, state: FSMContext):
    if not await check_channel_sub(message.from_user.id):
        await message.answer("⚠️ Avval kanalga obuna bo'ling:", reply_markup=get_subscription_keyboard())
        return

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
        await prepare_or_pay(call.message, call.from_user.id, state)

@dp.message(PresentationState.waiting_for_script)
async def process_script_input(message: Message, state: FSMContext):
    await state.update_data(user_script=message.text.strip())
    await prepare_or_pay(message, message.from_user.id, state)

async def prepare_or_pay(message: Message, user_id: int, state: FSMContext):
    data = await state.get_data()
    user_script = data.get("user_script")

    # Narxlar
    price_sum = 5000 if user_script else 7000

    # Tekin rejim yoki Faol obuna bo'lsa darhol yaratamiz
    if is_free_mode() or has_active_subscription(user_id):
        await generate_and_deliver(message, user_id, state)
    else:
        # Aks holda karta raqam va chek talab qilamiz
        await state.set_state(PaymentState.waiting_for_receipt)
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_payment")]
        ])
        await message.answer(
            f"💳 **To'lov qilish kerak:**\n\n"
            f"💰 Summa: **{price_sum:,} so'm**\n"
            f"💳 Karta: `{CARD_NUMBER}`\n"
            f"👤 Egasining ismi: **{CARD_HOLDER}**\n\n"
            f"⚠️ **Yo'riqnoma:** Yuqoridagi kartaga to'lovni amalga oshiring va **to'lov chekini (rasm/skrinshot)** ushbu chatga yuboring.\n\n"
            f"Adminlarimiz chekni ko'rib chiqqach, prezentatsiyangiz darhol tayyorlanadi!",
            reply_markup=kb,
            parse_mode="Markdown"
        )

@dp.callback_query(F.data == "cancel_payment")
async def cancel_payment_callback(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await call.message.answer("❌ To'lov bekor qilindi.", reply_markup=get_main_keyboard())

# CHEKNI QABUL QILISH VA ADMINGA YUBORISH
@dp.message(PaymentState.waiting_for_receipt, F.photo)
async def process_receipt_photo(message: Message, state: FSMContext):
    data = await state.get_data()
    topic = data.get("topic")
    user_script = data.get("user_script")
    price_sum = 5000 if user_script else 7000

    user_id = message.from_user.id
    full_name = message.from_user.full_name
    username = f"@{message.from_user.username}" if message.from_user.username else "Username yo'q"
    photo_id = message.photo[-1].file_id

    # Foydalanuvchiga bildirishnoma
    await message.answer("⏳ **Chekingiz qabul qilindi!**\nAdminlarimiz 1-15 daqiqa ichida tekshirib, prezentatsiyangizni yuborishadi.", parse_mode="Markdown")

    # Adminga chekni va tasdiqlash tugmalarini yuborish
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"approve_{user_id}"),
            InlineKeyboardButton(text="❌ Rad etish", callback_data=f"reject_{user_id}")
        ]
    ])

    try:
        await bot.send_photo(
            chat_id=ADMIN_ID,
            photo=photo_id,
            caption=(
                f"📥 **Yangi to'lov cheki!**\n\n"
                f"👤 Foydalanuvchi: {full_name} ({username})\n"
                f"🆔 ID: `{user_id}`\n"
                f"📝 Mavzu: {topic}\n"
                f"💵 Kutilayotgan summa: **{price_sum:,} so'm**"
            ),
            reply_markup=kb,
            parse_mode="Markdown"
        )
    except Exception as e:
        logging.error(f"Adminga chek yuborishda xatolik: {e}")

# ADMIN TASDIQLASHI YOKI RAD ETISHI
@dp.callback_query(F.data.startswith("approve_"))
async def approve_payment_callback(call: CallbackQuery, state: FSMContext):
    if call.from_user.id != ADMIN_ID:
        return

    target_user_id = int(call.data.replace("approve_", ""))
    await call.message.edit_caption(caption=call.message.caption + "\n\n✅ **TO'LOV TASDIQLANDI!**")

    # Foydalanuvchiga xabar berish va presentation yaratish
    try:
        await bot.send_message(chat_id=target_user_id, text="✅ **To'lovingiz tasdiqlandi!** Prezentatsiya tayyorlanmoqda...")
        await generate_and_deliver_direct(target_user_id, state)
    except Exception as e:
        await call.message.answer(f"❌ Xatolik yuz berdi: {e}")

@dp.callback_query(F.data.startswith("reject_"))
async def reject_payment_callback(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return

    target_user_id = int(call.data.replace("reject_", ""))
    await call.message.edit_caption(caption=call.message.caption + "\n\n❌ **TO'LOV RAD ETILDI!**")

    try:
        await bot.send_message(chat_id=target_user_id, text="❌ **To'lov chekingiz rad etildi.** Iltimos, qaytadan to'lov qiling yoki adminga murojaat qiling.")
    except Exception as e:
        logging.error(f"Foydalanuvchiga xabar yuborishda xatolik: {e}")

async def generate_and_deliver(message: Message, user_id: int, state: FSMContext):
    msg = await message.answer("⏳ Prezentatsiya va rasmlar tayyorlanmoqda, iltimos kuting...")
    await generate_and_deliver_direct(user_id, state)

async def generate_and_deliver_direct(user_id: int, state: FSMContext):
    data = await state.get_data()
    topic = data.get("topic")
    user_script = data.get("user_script")

    try:
        slides = await generate_presentation_content(topic, user_script)
        file_path = f"presentation_{user_id}.pptx"
        await create_pptx_file(slides, file_path)

        doc = FSInputFile(file_path)
        await bot.send_document(chat_id=user_id, document=doc, caption=f"✅ Prezentatsiya tayyor!\nMavzu: {topic}")
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        await bot.send_message(chat_id=user_id, text=f"❌ Xatolik yuz berdi: {e}")

    await state.clear()

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
