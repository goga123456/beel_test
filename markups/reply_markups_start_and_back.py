from aiogram.types import ReplyKeyboardMarkup, KeyboardButton


def get_start_and_back_kb() -> ReplyKeyboardMarkup:
    kmain = ReplyKeyboardMarkup(resize_keyboard=True)
    b1 = KeyboardButton('Назад')
    b2 = KeyboardButton('/start')
    kmain.add(b1, b2)
    return kmain

def get_start_kb() -> ReplyKeyboardMarkup:
    kmain2 = ReplyKeyboardMarkup(resize_keyboard=True)
    b1 = KeyboardButton('/start')
    kmain2.add(b1)
    return kmain2