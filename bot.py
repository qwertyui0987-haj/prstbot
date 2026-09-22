import os
import asyncio
import logging
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup, FSInputFile
from aiohttp import web
from dotenv import load_dotenv

from generator import generate_presentation_content, create_pptx_file

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
PORT = int(os.getenv("PORT", 8080))

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

class PresentationState(StatesGroup):
    waiting_for_script = State()
    waiting_for_topic = State()

def get_main_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Tayyor ssenariy bilan (15 000 so'm)", callback_data="plan_standard")],
        [InlineKeyboardButton(text="🤖 AI avtomatik yaratishi (20 000 so'm)", callback_data="plan_premium")]
    ])

@dp.message(CommandStart())
async def start_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        f"Assalomu alaykum, {message.from_user.full_name}!\n\n"
        "Prezentatsiya yaratuvchi botga xush kelibsiz.\n"
        "Kerakli tarifni tanlang:",
        reply_markup=get_main_keyboard()
    )

@dp.callback_query(F.data == "plan_standard")
async def process_standard_plan(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(PresentationState.waiting_for_script)
    await callback.message.answer(
        "Siz 15 000 so'mlik tarifni tanladingiz.\n\n"
        "Iltimos, prezentatsiya mavzusini va qisqacha ssenariyingizni yuboring:"
    )
    await callback.answer()

@dp.callback_query(F.data == "plan_premium")
async def process_premium_plan(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(PresentationState.waiting_for_topic)
    await callback.message.answer(
        "Siz 20 000 so'mlik tarifni tanladingiz.\n\n"
        "Iltimos, faqat prezentatsiya mavzusini yuboring (Masalan: Sun'iy intellektning kelajagi):"
    )
    await callback.answer()

@dp.message(PresentationState.waiting_for_script)
async def handle_standard(message: types.Message, state: FSMContext):
    msg = await message.answer("⏳ Gemini AI ssenariyingizni qayta ishlamoqda...")
    try:
        topic = message.text[:50]
        slides_data = await generate_presentation_content(topic=topic, user_script=message.text)
        
        file_path = f"pres_{message.from_user.id}.pptx"
        await create_pptx_file(slides_data, file_path)
        
        doc = FSInputFile(file_path)
        await message.answer_document(doc, caption="✅ Prezentatsiyangiz tayyor!")
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        logging.exception(f"XATOLIK YUZ BERDI: {e}")
        await message.answer(f"❌ Prezentatsiya yaratishda xatolik yuz berdi:\n{e}")
    finally:
        try:
            await msg.delete()
        except Exception:
            pass
        await state.clear()

@dp.message(PresentationState.waiting_for_topic)
async def handle_premium(message: types.Message, state: FSMContext):
    msg = await message.answer("⏳ Gemini AI mavzu bo'yicha slaydlar matnini tuzmoqda...")
    try:
        slides_data = await generate_presentation_content(topic=message.text)
        
        file_path = f"pres_{message.from_user.id}.pptx"
        await create_pptx_file(slides_data, file_path)
        
        doc = FSInputFile(file_path)
        await message.answer_document(doc, caption="✅ AI tomonidan yaratilgan prezentatsiyangiz tayyor!")
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        logging.exception(f"XATOLIK YUZ BERDI: {e}")
        await message.answer(f"❌ Prezentatsiya yaratishda xatolik yuz berdi:\n{e}")
    finally:
        try:
            await msg.delete()
        except Exception:
            pass
        await state.clear()

# --- Oddiy matn yuborilganda ham ishlaydigan Handler (Tugma bosilmagan bo'lsa ham) ---
@dp.message(F.text)
async def handle_direct_text(message: types.Message, state: FSMContext):
    await state.set_state(PresentationState.waiting_for_topic)
    await handle_premium(message, state)

# --- Render WebService uchun HTTP Server ---
async def handle_ping(request):
    return web.Response(text="Bot is running active 24/7!")

async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logging.info(f"Web server started on port {PORT}")

async def main():
    await start_web_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
