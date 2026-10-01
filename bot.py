import os
import sys
import sqlite3
import asyncio
import datetime
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    LabeledPrice,
    PreCheckoutQuery,
    FSInputFile
)

# generator.py funksiyalarini ulash
from generator import generate_presentation_content, create_pptx_file

# ================= TOKEN VA SOZLAMALAR =================
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
ADMIN_ID = int(os.getenv("ADMIN_ID", "123456789"))  # Admin Telegram ID si

# O'zbekistonda eng qulayi Telegram Payments (Click / Payme provider token yoki Telegram Stars "")
PAYMENT_PROVIDER_TOKEN = os.getenv("PAYMENT_PROVIDER_TOKEN", "") # BotFather'dan olingan Click/Payme token
SLIDE_PRICE_SUM = 15000  # Bitta prezentatsiya narxi (so'mda)

# Majburiy obuna kanali username'i (Botingiz ushbu kanalda ADMIN bo'lishi shart!)
REQUIRED_CHANNEL = os.getenv("REQUIRED_CHANNEL", "@sizning_kanalingiz") 

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# ================= MA'LUMOTLAR BAZASI (SQLite) =================
DB_NAME = "bot_data.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Foydalanuvchilar jadvali
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            sub_until INTEGER DEFAULT 0, -- UNIX timestamp (-1 = umrbod)
            joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Bot sozlamalari (Masalan: tekin rejim yoqilganmi)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    
    # Promokodlar jadvali
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS promocodes (
            code TEXT PRIMARY KEY,
            duration_type TEXT, -- 'unlimited', '1_month', '1_week', '1_day'
            is_used INTEGER DEFAULT 0,
            used_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Statistika jadvali
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stats (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            total_slides INTEGER DEFAULT 0
        )
    """)
    
    # Boshlang'ich qiymatlar
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('free_mode', '0')") # 0 = To'lovli, 1 = Tekin
    cursor.execute("INSERT OR IGNORE INTO stats (id, total_slides) VALUES (1, 0)")
    
    conn.commit()
    conn.close()

init_db()

# Baza bilan ishlash uchun yordamchi funksiyalar
def get_setting(key: str, default: str = "0") -> str:
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else default

def set_setting(key: str, value: str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def add_or_update_user(user_id: int, username: str, full_name: str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO users (user_id, username, full_name)
        VALUES (?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET username=?, full_name=?
    """, (user_id, username, full_name, username, full_name))
    conn.commit()
    conn.close()

def is_user_subscribed(user_id: int) -> bool:
    # 1. Agar tekin rejim yoqilgan bo'lsa -> Hamma ishlatishi mumkin
    if get_setting("free_mode") == "1":
        return True
    
    # 2. Obuna vaqtini Kerakli timestamp bilan solishtirish
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT sub_until FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        return False
    
    sub_until = row[0]
    if sub_until == -1: # Umrbod tekin
        return True
    
    now_ts = int(datetime.datetime.now().timestamp())
    return sub_until > now_ts

def increment_slide_count():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("UPDATE stats SET total_slides = total_slides + 1 WHERE id = 1")
    conn.commit()
    conn.close()

# ================= MAJBURIY OBUNA TEKSHIRUV YoRDAMChISI =================
async def is_subscribed_to_channel(user_id: int) -> bool:
    if not REQUIRED_CHANNEL or REQUIRED_CHANNEL == "@sizning_kanalingiz":
        return True  # Agar kanal username sozlangan bo'lmasa, o'tkazib yuboradi
    try:
        member = await bot.get_chat_member(chat_id=REQUIRED_CHANNEL, user_id=user_id)
        return member.status in ["creator", "administrator", "member"]
    except Exception as e:
        print(f"[SUB CHECK ERROR] Kanal tekshirishda xatolik: {e}")
        return True # Bot kanalda admin bo'lmasa xatolik bermasligi uchun

def get_sub_keyboard():
    channel_url = f"https://t.me/{REQUIRED_CHANNEL.replace('@', '')}"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📢 Kanalga obuna bo'lish", url=channel_url)],
        [InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_subscription")]
    ])
    return keyboard

# ================= FSM HOLATLARI =================
class SlideForm(StatesGroup):
    waiting_for_topic = State()
    waiting_for_script = State()

class AdminForm(StatesGroup):
    waiting_for_promo_code = State()
    waiting_for_broadcast = State()

# ================= TUGMALAR (KEYBOARDS) =================
def get_main_keyboard(user_id: int):
    buttons = [
        [KeyboardButton(text="📊 Slayd yaratish"), KeyboardButton(text="🎟 Promokod kiritish")],
        [KeyboardButton(text="💳 Obuna holati")]
    ]
    if user_id == ADMIN_ID:
        buttons.append([KeyboardButton(text="⚙️ Admin Panel")])
    
    return types.ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

def get_admin_inline_keyboard():
    free_mode = get_setting("free_mode") == "1"
    mode_text = "🟢 Rejim: TEKIN (O'chirish)" if free_mode else "🔴 Rejim: TO'LOVLI (Tekin qilish)"
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=mode_text, callback_data="toggle_free_mode")],
        [InlineKeyboardButton(text="➕ Yangi Promokod yaratish", callback_data="create_promo")],
        [InlineKeyboardButton(text="📈 Statistika", callback_data="admin_stats")],
        [InlineKeyboardButton(text="📢 Xabar yuborish (Broadcast)", callback_data="admin_broadcast")]
    ])
    return keyboard

# ================= COMMAND HANDLERS =================

@dp.message(CommandStart())
async def cmd_start(message: Message):
    add_or_update_user(message.from_user.id, message.from_user.username, message.from_user.full_name)
    
    # Kanalga obunani tekshirish
    if not await is_subscribed_to_channel(message.from_user.id):
        await message.answer(
            f"⚠️ Botdan foydalanish uchun rasmiy kanalimizga obuna bo'ling:",
            reply_markup=get_sub_keyboard()
        )
        return

    free_mode = get_setting("free_mode") == "1"
    welcome_msg = (
        f"Assalomu alaykum, {message.from_user.first_name}!\n\n"
        f"🤖 Men sun'iy intellekt yordamida daqiqalar ichida **peffekt prezentatsiyalar** yaratib beruvchi botman.\n\n"
    )
    if free_mode:
        welcome_msg += "🎉 **Hozirda bot aksiyada! Hamma uchun TEKIN rejim yoqilgan!**"
    else:
        welcome_msg += "💡 Slayd yaratish uchun mavzu yuboring yoki menyudan foydalaning."

    await message.answer(welcome_msg, reply_markup=get_main_keyboard(message.from_user.id), parse_mode="Markdown")

@dp.callback_query(F.data == "check_subscription")
async def check_sub_callback(callback: CallbackQuery):
    if await is_subscribed_to_channel(callback.from_user.id):
        await callback.message.delete()
        await callback.message.answer(
            "✅ Obunangiz tasdiqlandi! Endi botdan to'liq foydalanishingiz mumkin.",
            reply_markup=get_main_keyboard(callback.from_user.id)
        )
    else:
        await callback.answer("❌ Siz hali kanalga obuna bo'lmadingiz!", show_alert=True)

# ================= PROMAKOD BOSHQRUVI =================

@dp.message(F.text == "🎟 Promokod kiritish")
async def ask_promo(message: Message, state: FSMContext):
    await message.answer("🔑 Promokodingizni yuboring:")
    await state.set_state("waiting_for_user_promo")

@dp.message(F.state == "waiting_for_user_promo")
async def process_user_promo(message: Message, state: FSMContext):
    code = message.text.strip()
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("SELECT duration_type, is_used FROM promocodes WHERE code = ?", (code,))
    row = cursor.fetchone()
    
    if not row:
        await message.answer("❌ Bunday promokod topilmadi yoki xato kiritildi.")
        await state.clear()
        conn.close()
        return
    
    duration_type, is_used = row
    if is_used == 1:
        await message.answer("⚠️ Bu promokod allaqachon ishlatilgan.")
        await state.clear()
        conn.close()
        return
    
    # Vaqtni hisoblash
    now = datetime.datetime.now()
    if duration_type == "unlimited":
        sub_until = -1
        msg_type = "Umrbod tekin foydalanish"
    elif duration_type == "1_month":
        sub_until = int((now + datetime.timedelta(days=30)).timestamp())
        msg_type = "1 Oylik tekin foydalanish"
    elif duration_type == "1_week":
        sub_until = int((now + datetime.timedelta(days=7)).timestamp())
        msg_type = "1 Haftalik tekin foydalanish"
    elif duration_type == "1_day":
        sub_until = int((now + datetime.timedelta(days=1)).timestamp())
        msg_type = "1 Kunlik tekin foydalanish"
    else:
        sub_until = 0
        msg_type = "Noma'lum"

    # Promokodni aktivlashtirish
    cursor.execute("UPDATE promocodes SET is_used = 1, used_by = ? WHERE code = ?", (message.from_user.id, code))
    cursor.execute("UPDATE users SET sub_until = ? WHERE user_id = ?", (sub_until, message.from_user.id))
    conn.commit()
    conn.close()
    
    await message.answer(f"🎉 Tabriklaymiz! **{code}** promokodi muvaffaqiyatli faollashtirildi!\n🎁 Sizga: **{msg_type}** berildi.")
    await state.clear()

# ================= ADMIN PANEL =================

@dp.message(F.text == "⚙️ Admin Panel")
@dp.message(Command("admin"))
async def admin_panel(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    
    await message.answer("🛠 **Admin Boshqaruv Paneli:**", reply_markup=get_admin_inline_keyboard(), parse_mode="Markdown")

@dp.callback_query(F.data == "toggle_free_mode")
async def toggle_free_mode_callback(callback: CallbackQuery):
    if callback.from_user.id != ADMIN_ID:
        return
    
    current = get_setting("free_mode")
    new_status = "0" if current == "1" else "1"
    set_setting("free_mode", new_status)
    
    status_text = "🟢 **Bot TEKIN rejimga o'tkazildi!** Endi hamma tekin slayd yarata oladi." if new_status == "1" else "🔴 **Bot TO'LOVLI rejimga o'tkazildi!**"
    await callback.message.edit_text(f"✅ Rejim o'zgartirildi!\n\n{status_text}", reply_markup=get_admin_inline_keyboard(), parse_mode="Markdown")

@dp.callback_query(F.data == "admin_stats")
async def admin_stats_callback(callback: CallbackQuery):
    if callback.from_user.id != ADMIN_ID:
        return
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    
    cursor.execute("SELECT total_slides FROM stats WHERE id = 1")
    total_slides = cursor.fetchone()[0]
    
    free_mode = "YOQILGAN (Tekin)" if get_setting("free_mode") == "1" else "O'CHIRILGAN (To'lovli)"
    conn.close()
    
    text = (
        f"📊 **BOT STATISTIKASI:**\n\n"
        f"👥 Jami foydalanuvchilar: **{total_users} ta**\n"
        f"📑 Yaratilgan slaydlar: **{total_slides} ta**\n"
        f"⚡️ Bosh rejim: **{free_mode}**"
    )
    await callback.message.answer(text, parse_mode="Markdown")
    await callback.answer()

@dp.callback_query(F.data == "create_promo")
async def admin_create_promo_start(callback: CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_ID:
        return
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="♾ Umrbod Tekin", callback_data="promo_type:unlimited")],
        [InlineKeyboardButton(text="📅 1 Oylik", callback_data="promo_type:1_month")],
        [InlineKeyboardButton(text="📆 1 Haftalik", callback_data="promo_type:1_week")],
        [InlineKeyboardButton(text="⏳ 1 Kunlik", callback_data="promo_type:1_day")]
    ])
    await callback.message.answer("🏷 Promokod muddatini tanlang:", reply_markup=keyboard)
    await callback.answer()

@dp.callback_query(F.data.startswith("promo_type:"))
async def admin_promo_type_chosen(callback: CallbackQuery, state: FSMContext):
    promo_type = callback.data.split(":")[1]
    await state.update_data(promo_type=promo_type)
    
    await callback.message.answer("✍️ Promokod so'zini yuboring (masalan: `YANGI2026` yoki `SPECIAL`):")
    await state.set_state(AdminForm.waiting_for_promo_code)
    await callback.answer()

@dp.message(AdminForm.waiting_for_promo_code)
async def admin_save_promo(message: Message, state: FSMContext):
    promo_code = message.text.strip().upper()
    data = await state.get_data()
    promo_type = data.get("promo_type", "1_month")
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO promocodes (code, duration_type) VALUES (?, ?)", (promo_code, promo_type))
        conn.commit()
        conn.close()
        
        await message.answer(f"✅ Promokod yaratildi:\n\n🔑 Kod: `{promo_code}`\n⏳ Muxlati: **{promo_type}**", parse_mode="Markdown")
    except sqlite3.IntegrityError:
        await message.answer("⚠️️ Bu nomdagi promokod allaqachon mavjud! Boshqa nom kiritib ko'ring.")
    
    await state.clear()

# ================= SLAYD YARATISH VA TO'LOV TIZIMI =================

@dp.message(F.text == "📊 Slayd yaratish")
async def start_slide_creation(message: Message, state: FSMContext):
    user_id = message.from_user.id
    
    # 0. Avval majburiy kanal obunasini tekshirish
    if not await is_subscribed_to_channel(user_id):
        await message.answer(
            "⚠️ Slayd yaratish uchun avval kanalimizga obuna bo'lishingiz kerak:",
            reply_markup=get_sub_keyboard()
        )
        return

    # 1. Obunani tekshirish
    if is_user_subscribed(user_id):
        await message.answer("✍️ Prezentatsiya mavzusini kiriting:")
        await state.set_state(SlideForm.waiting_for_topic)
    else:
        # Obunasi yo'q bo'lsa -> To'lov schyotini (Invoice) yuborish
        await send_payment_invoice(message)

async def send_payment_invoice(message: Message):
    """Click / Payme / Telegram Stars orqali avtomatik to'lov taklif qilish"""
    
    if not PAYMENT_PROVIDER_TOKEN:
        # Agar to'lov tokeni qo'yilmagan bo'lsa
        await message.answer(
            "💳 **Slayd yaratish uchun to'lov qilish lozim.**\n\n"
            f"💰 Bitta slayd narxi: **{SLIDE_PRICE_SUM:,} so'm**\n\n"
            "*(Eslatma: Bot administratoriga murojaat qiling yoki Promokod kiriting)*"
        )
        return

    prices = [LabeledPrice(label="AI Prezentatsiya Yaratish", amount=SLIDE_PRICE_SUM * 100)] # tiyinlarda
    
    await bot.send_invoice(
        chat_id=message.chat.id,
        title="📑 AI Slayd Prezentatsiya",
        description="Sun'iy intellekt yordamida 6 ta slaydli professional prezentatsiya yaratish xizmati.",
        provider_token=PAYMENT_PROVIDER_TOKEN,
        currency="UZS",
        prices=prices,
        start_parameter="slide-presentation-payment",
        payload=f"user_pay_{message.from_user.id}"
    )

# Pre-checkout check (Telegram Payments talabi)
@dp.pre_checkout_query()
async def process_pre_checkout(pre_checkout_query: PreCheckoutQuery):
    await bot.answer_pre_checkout_query(pre_checkout_query.id, ok=True)

# Muvaffaqiyatli to'lov amalga oshgach
@dp.message(F.successful_payment)
async def process_successful_payment(message: Message, state: FSMContext):
    await message.answer("🎉 **To'lov muvaffaqiyatli amalga oshirildi!**\n\nEndi prezentatsiya mavzusini yuboring:")
    await state.set_state(SlideForm.waiting_for_topic)

# ================= GENERATION PROCESS =================

@dp.message(SlideForm.waiting_for_topic)
async def process_topic(message: Message, state: FSMContext):
    topic = message.text.strip()
    
    msg = await message.answer("🧠 Sun'iy intellekt slayd mazmunini shakllantirmoqda va rasmlarni tanlamoqda... (10-20 soniya kutib turing)")
    
    try:
        # 1. AI orqali tarkib yaratish
        slides_data = await generate_presentation_content(topic=topic)
        
        await msg.edit_text("🎨 PowerPoint (.pptx) fayli hosil qilinmoqda...")
        
        # 2. PPTX faylni yaratish
        output_filename = f"presentation_{message.from_user.id}.pptx"
        file_path = await create_pptx_file(slides_data, output_filename)
        
        # 3. Faylni foydalanuvchiga yuborish
        await msg.delete()
        pptx_file = FSInputFile(file_path)
        await message.answer_document(
            document=pptx_file,
            caption=f"✅ **Sizning prezentatsiyangiz tayyor!**\n\n📌 Mavzu: **{topic}**\n📄 Slaydlar soni: 6 ta"
        )
        
        # Statistika yangilash
        increment_slide_count()
        
        # Faylni o'chirish
        if os.path.exists(file_path):
            os.remove(file_path)
            
    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {e}")
    
    await state.clear()

# ================= BOTNI ISHGA TUSHIRISH =================
async def main():
    print("[BOT LOG] Bot muvaffaqiyatli ishga tushdi!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
```

### Muhim eslatma:
1. `REQUIRED_CHANNEL` o'rniga o'zingizning kanalingiz usernamesini yozing (masalan `@my_channel_name`).
2. Telegram botingizni kanalingizga **Admin** qilib qo'shing, aks holda Telegram API a'zolarni tekshirishga ruxsat bermaydi.
