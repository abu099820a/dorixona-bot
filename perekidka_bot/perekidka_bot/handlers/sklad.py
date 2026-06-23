from aiogram import Router, F
from aiogram.types import Message, CallbackQuery

from database import get_user, get_orders, update_order_status
from keyboards.main import order_actions
from utils.sheets import update_sheet_status

router = Router()

@router.message(F.text == "📤 Жўнатиш")
async def send_items(message: Message):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "sklad":
        return

    orders = await get_orders(status="yangi")
    if not orders:
        await message.answer("📭 Жўнатиладиган буюртма йўқ.")
        return

    for o in orders:
        text = (
            f"🆕 Буюртма #{o['id']}\n"
            f"📍 {o['from_apteka']} → {o['to_apteka']}\n"
            f"💊 {o['tovar']} — {o['miqdor']}"
        )
        await message.answer(text, reply_markup=order_actions(o["id"], "sklad"))

@router.callback_query(F.data.startswith("sent_"))
async def mark_sent(callback: CallbackQuery):
    order_id = int(callback.data.split("_")[1])
    user = await get_user(callback.from_user.id)
    if not user or user["role"] != "sklad":
        return

    await update_order_status(order_id, "jarayonda")
    await update_sheet_status(order_id, "jarayonda")
    await callback.message.edit_text(
        callback.message.text + "\n\n📤 Жўнатилди!"
    )
    await callback.answer("✅ Белгиланди!")

@router.message(F.text == "📋 Барча буюртмалар")
async def all_orders(message: Message):
    user = await get_user(message.from_user.id)
    if not user or user["role"] != "sklad":
        return

    orders = await get_orders()
    if not orders:
        await message.answer("📭 Буюртма йўқ.")
        return

    text = "📋 Барча буюртмалар:\n\n"
    status_map = {"yangi": "🆕 Янги", "jarayonda": "🚗 Йўлда",
                  "yetkazildi": "✔️ Етказилди", "qabul": "📥 Қабул"}
    for o in orders[:15]:
        st = status_map.get(o["status"], o["status"])
        text += f"{st} | #{o['id']} {o['from_apteka']}→{o['to_apteka']} | {o['tovar']}\n"

    await message.answer(text)
