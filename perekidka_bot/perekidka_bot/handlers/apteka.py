from aiogram import Router, F
from aiogram.types import Message, CallbackQuery

from database import get_user, get_orders, update_order_status
from keyboards.main import order_actions
from utils.sheets import update_sheet_status

router = Router()

@router.message(F.text == "📥 Кутилаётган товарлар")
async def waiting_items(message: Message):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "apteka":
        return

    orders = await get_orders(status="yetkazildi")
    if not orders:
        await message.answer("📭 Кутилаётган товар йўқ.")
        return

    for o in orders:
        text = (
            f"🚗 Йўлда — Буюртма #{o['id']}\n"
            f"📍 {o['from_apteka']} дан\n"
            f"💊 {o['tovar']} — {o['miqdor']}"
        )
        await message.answer(text, reply_markup=order_actions(o["id"], "apteka"))

@router.callback_query(F.data.startswith("received_"))
async def received_order(callback: CallbackQuery):
    order_id = int(callback.data.split("_")[1])
    user = await get_user(callback.from_user.id)
    if not user or user["role"] != "apteka":
        return

    await update_order_status(order_id, "qabul")
    await update_sheet_status(order_id, "qabul")
    await callback.message.edit_text(
        callback.message.text + "\n\n📥 Қабул қилинди!"
    )
    await callback.answer("✅ Товар қабул қилинди!")
