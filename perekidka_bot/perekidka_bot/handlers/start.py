from aiogram import Router, F
from aiogram.types import Message
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from config import PASSWORDS, ROLE_NAMES
from database import save_user, get_user
from keyboards.main import main_menu

router = Router()

class AuthStates(StatesGroup):
    waiting_password = State()

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    user = await get_user(message.from_user.id)
    if user:
        role = user["role"]
        await message.answer(
            f"👋 Хуш келибсиз, {ROLE_NAMES[role]}!\n"
            f"Менюдан танланг:",
            reply_markup=main_menu(role)
        )
        return

    await message.answer(
        "👋 Перекидка ботига хуш келибсиз!\n\n"
        "🔐 Паролингизни киритинг:"
    )
    await state.set_state(AuthStates.waiting_password)

@router.message(AuthStates.waiting_password)
async def check_password(message: Message, state: FSMContext):
    password = message.text.strip()
    role = None

    for r, p in PASSWORDS.items():
        if p == password:
            role = r
            break

    if not role:
        await message.answer("❌ Нотўғри парол. Қайтадан уриниб кўринг:")
        return

    await save_user(message.from_user.id, message.from_user.username, role)
    await state.clear()

    await message.answer(
        f"✅ Хуш келибсиз, {ROLE_NAMES[role]}!\n"
        f"Менюдан танланг:",
        reply_markup=main_menu(role)
    )

@router.message(F.text == "🚪 Чиқиш")
async def logout(message: Message, state: FSMContext):
    from database import DB_PATH
    import aiosqlite
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("DELETE FROM users WHERE user_id=?", (message.from_user.id,))
        await db.commit()

    await state.clear()
    await message.answer(
        "👋 Чиқдингиз. Қайта кириш учун /start босинг.",
        reply_markup=None
    )
