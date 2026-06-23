from aiogram import Router, F
from aiogram.types import Message
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database import get_user, create_order, get_orders
from utils.sheets import log_to_sheets

router = Router()

class OrderStates(StatesGroup):
    from_apteka = State()
    to_apteka   = State()
    tovar       = State()
    miqdor      = State()

@router.message(F.text == "📦 Янги буюртма")
async def new_order_start(message: Message, state: FSMContext):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "snabjenets":
        return

    await message.answer("📍 Қайси аптекадан жўнатилади? (номини ёзинг)")
    await state.set_state(OrderStates.from_apteka)

@router.message(OrderStates.from_apteka)
async def order_from(message: Message, state: FSMContext):
    await state.update_data(from_apteka=message.text)
    await message.answer("📍 Қайси аптекага жўнатилади?")
    await state.set_state(OrderStates.to_apteka)

@router.message(OrderStates.to_apteka)
async def order_to(message: Message, state: FSMContext):
    await state.update_data(to_apteka=message.text)
    await message.answer("💊 Товар номи нима?")
    await state.set_state(OrderStates.tovar)

@router.message(OrderStates.tovar)
async def order_tovar(message: Message, state: FSMContext):
    await state.update_data(tovar=message.text)
    await message.answer("🔢 Миқдори қанча? (dona, kg, упак)")
    await state.set_state(OrderStates.miqdor)

@router.message(OrderStates.miqdor)
async def order_miqdor(message: Message, state: FSMContext):
    data = await state.get_data()
    await state.clear()

    order_id = await create_order(
        from_apteka=data["from_apteka"],
        to_apteka=data["to_apteka"],
        tovar=data["tovar"],
        miqdor=message.text,
        created_by=message.from_user.id
    )

    await log_to_sheets(order_id, data["from_apteka"], data["to_apteka"],
                        data["tovar"], message.text, "yangi")

    await message.answer(
        f"✅ Буюртма #{order_id} яратилди!\n\n"
        f"📍 {data['from_apteka']} → {data['to_apteka']}\n"
        f"💊 {data['tovar']} — {message.text}\n"
        f"📊 Статус: Янги"
    )

@router.message(F.text == "📋 Менинг буюртмаларим")
async def my_orders(message: Message, state: FSMContext):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "snabjenets":
        return

    orders = await get_orders()
    if not orders:
        await message.answer("📭 Ҳозирча буюртма йўқ.")
        return

    text = "📋 Сўнгги буюртмалар:\n\n"
    for o in orders[:10]:
        status_emoji = {"yangi": "🆕", "jarayonda": "🚗", "yetkazildi": "✅", "qabul": "📥"}.get(o["status"], "❓")
        text += f"{status_emoji} #{o['id']} | {o['from_apteka']} → {o['to_apteka']}\n"
        text += f"   💊 {o['tovar']} ({o['miqdor']})\n\n"

    await message.answer(text)
