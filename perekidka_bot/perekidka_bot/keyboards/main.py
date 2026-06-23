from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton

def main_menu(role: str):
    if role == "snabjenets":
        buttons = [
            [KeyboardButton(text="📦 Янги буюртма")],
            [KeyboardButton(text="📋 Менинг буюртмаларим")],
            [KeyboardButton(text="📊 Статистика")],
        ]
    elif role == "shofer":
        buttons = [
            [KeyboardButton(text="🚗 Янги топшириқлар")],
            [KeyboardButton(text="✅ Қабул қилдим")],
            [KeyboardButton(text="📋 Менинг тарихим")],
        ]
    elif role == "apteka":
        buttons = [
            [KeyboardButton(text="📥 Кутилаётган товарлар")],
            [KeyboardButton(text="✅ Товарни қабул қилдим")],
            [KeyboardButton(text="📋 Тарих")],
        ]
    elif role == "sklad":
        buttons = [
            [KeyboardButton(text="📤 Жўнатиш")],
            [KeyboardButton(text="📋 Барча буюртмалар")],
            [KeyboardButton(text="📊 Ҳисобот")],
        ]
    else:
        buttons = []

    buttons.append([KeyboardButton(text="🚪 Чиқиш")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

def order_actions(order_id: int, role: str):
    buttons = []
    if role == "shofer":
        buttons = [
            [InlineKeyboardButton(text="✅ Қабул қилдим", callback_data=f"take_{order_id}")],
            [InlineKeyboardButton(text="✔️ Етказдим",     callback_data=f"done_{order_id}")],
        ]
    elif role == "apteka":
        buttons = [
            [InlineKeyboardButton(text="✅ Олдим", callback_data=f"received_{order_id}")],
        ]
    elif role == "sklad":
        buttons = [
            [InlineKeyboardButton(text="📤 Жўнатилди", callback_data=f"sent_{order_id}")],
        ]
    return InlineKeyboardMarkup(inline_keyboard=buttons) if buttons else None
