import asyncio
import logging
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup

# Bot tokeningizni kiriting
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Foydalanuvchi holatlari (FSM)
class PresentationState(StatesGroup):
    choosing_plan = State()
    waiting_for_topic = State()
    waiting_for_script = State()

# Inline tugmalar
def get_main_keyboard():
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Tayyor ssenariy bilan (15 000 so'm)", callback_data="plan_standard")],
        [InlineKeyboardButton(text="🤖 AI avtomatik yaratishi (20 000 so'm)", callback_data="plan_premium")]
    ])
    return keyboard

@dp.message(CommandStart())
async def start_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        f"Assalomu alaykum, {message.from_user.full_name}!\n\n"
        "Prezentatsiya yaratuvchi botga xush kelibsiz.\n"
        "Iltimos, o'zingizga mos tarifni tanlang:",
        reply_markup=get_main_keyboard()
    )

@dp.callback_query(F.data == "plan_standard")
async def process_standard_plan(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(plan="standard")
    await state.set_state(PresentationState.waiting_for_script)
    await callback.message.answer(
        "Siz **15 000 so'mlik** tarifni tanladingiz.\n\n"
        "Iltimos, prezentatsiya mavzusi va slaydlar ssenariysini (matnlarini) yuboring."
    )
    await callback.answer()

@dp.callback_query(F.data == "plan_premium")
async def process_premium_plan(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(plan="premium")
    await state.set_state(PresentationState.waiting_for_topic)
    await callback.message.answer(
        "Siz **20 000 so'mlik** tarifni tanladingiz.\n\n"
        "Iltimos, prezentatsiya mavzusini yuboring. AI o'zi slaydlar matnini tuzib beradi."
    )
    await callback.answer()

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())