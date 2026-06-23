from aiogram import Router, F
from aiogram.types import Message, CallbackQuery

from database import get_user, get_orders, update_order_status
from keyboards.main import order_actions
from utils.sheets import update_sheet_status

router = Router()

@router.message(F.text == "🚗 Янги топшириқлар")
async def new_tasks(message: Message):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "shofer":
        return

    orders = await get_orders(status="yangi")
    if not orders:
        await message.answer("✅ Ҳозирча янги топшириқ йўқ.")
        return

    for o in orders:
        text = (
            f"🆕 Буюртма #{o['id']}\n"
            f"📍 {o['from_apteka']} → {o['to_apteka']}\n"
            f"💊 {o['tovar']} — {o['miqdor']}"
        )
        await message.answer(text, reply_markup=order_actions(o["id"], "shofer"))

@router.callback_query(F.data.startswith("take_"))
async def take_order(callback: CallbackQuery):
    order_id = int(callback.data.split("_")[1])
    user = await get_user(callback.from_user.id)
    if not user or user["role"] != "shofer":
        return

    await update_order_status(order_id, "jarayonda", shofer_id=callback.from_user.id)
    await update_sheet_status(order_id, "jarayonda")
    await callback.message.edit_text(
        callback.message.text + "\n\n🚗 Қабул қилдингиз!"
    )
    await callback.answer("✅ Топшириқ қабул қилинди!")

@router.callback_query(F.data.startswith("done_"))
async def done_order(callback: CallbackQuery):
    order_id = int(callback.data.split("_")[1])
    user = await get_user(callback.from_user.id)
    if not user or user["role"] != "shofer":
        return

    await update_order_status(order_id, "yetkazildi")
    await update_sheet_status(order_id, "yetkazildi")
    await callback.message.edit_text(
        callback.message.text + "\n\n✔️ Етказилди!"
    )
    await callback.answer("✅ Белгиланди!")
