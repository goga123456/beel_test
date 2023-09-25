from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

from messages import *


def get_initial_kb() -> InlineKeyboardMarkup:
    kb1 = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton('Продолжить', callback_data='next')
    b2 = InlineKeyboardButton('Отказаться', callback_data='close')
    kb1.add(b1, b2)
    return kb1


def get_initial_kb2() -> InlineKeyboardMarkup:
    kb2 = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton('Да,я хочу в билайн', callback_data='yes_i_want')
    b2 = InlineKeyboardButton('Не интересует', callback_data='i_dont_want')
    kb2.add(b1, b2)
    return kb2


def get_town_kb() -> InlineKeyboardMarkup:
    kb_town = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton('Ташкент', callback_data='Ташкент')
    b2 = InlineKeyboardButton('Другой город или регион', callback_data='Другой')
    b3 = InlineKeyboardButton('Назад', callback_data='back_to_birth')
    kb_town.add(b1, b2).add(b3)
    return kb_town


def get_district_kb() -> InlineKeyboardMarkup:
    kb_dist = InlineKeyboardMarkup(resize_keyboard=True, row_width=3)
    item1 = InlineKeyboardButton('Алмазар', callback_data='Алмазар')
    item2 = InlineKeyboardButton('Бектемир', callback_data='Бектемир')
    item3 = InlineKeyboardButton('Мирабад', callback_data='Мирабад')
    item4 = InlineKeyboardButton('Мирзо-Улугбек', callback_data='Мирзо-Улугбек')
    item5 = InlineKeyboardButton('Сергели', callback_data='Сергели')
    item6 = InlineKeyboardButton('Чиланзар', callback_data='Чиланзар')
    item7 = InlineKeyboardButton('Шайхантаур', callback_data='Шайхантаур')
    item8 = InlineKeyboardButton('Юнусабад', callback_data='Юнусабад')
    item9 = InlineKeyboardButton('Яккасарай', callback_data='Яккасарай')
    item10 = InlineKeyboardButton('Яшнабад', callback_data='Яшнабад')
    item11 = InlineKeyboardButton('Учтепа', callback_data='Учтепа')
    item12 = InlineKeyboardButton('Назад', callback_data='back_to_town')
    kb_dist.add(item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12)
    return kb_dist


def get_edu_kb() -> InlineKeyboardMarkup:
    kb_edu = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton('Высшее', callback_data='Высшее')
    b2 = InlineKeyboardButton('Неполное высшее', callback_data='Неполное высшее')
    b3 = InlineKeyboardButton('Среднее', callback_data='Среднее')
    b4 = InlineKeyboardButton('Неполное среднее', callback_data='Неполное среднее')
    b5 = InlineKeyboardButton('Среднее специальное', callback_data='Среднее специальное')
    b6 = InlineKeyboardButton('Назад', callback_data='to_town')
    kb_edu.add(b1).add(b2).add(b3).add(b4).add(b5).add(b6)
    return kb_edu


def get_rus_kb() -> InlineKeyboardMarkup:
    kb_rus = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton(text=great, callback_data='Отлично')
    b2 = InlineKeyboardButton(text=good, callback_data='Хорошо')
    b3 = InlineKeyboardButton(text=not_good, callback_data='Удовлетворительно')
    b4 = InlineKeyboardButton(text=bad, callback_data='Не владею русским языком')
    b5 = InlineKeyboardButton('Назад', callback_data='back_to_edu')
    kb_rus.add(b1).add(b2).add(b3).add(b4).add(b5)
    return kb_rus


def get_uzb_kb() -> InlineKeyboardMarkup:
    kb_uzb = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton(text=great, callback_data='Отлично знаю')
    b2 = InlineKeyboardButton(text=good, callback_data='Хорошо знаю')
    b3 = InlineKeyboardButton(text=not_good, callback_data='Удовлетворительно знаю')
    b4 = InlineKeyboardButton(text=bad, callback_data='Не владею узбекским языком')
    b5 = InlineKeyboardButton('Назад', callback_data='back_to_ru')
    kb_uzb.add(b1).add(b2).add(b3).add(b4).add(b5)
    return kb_uzb


def get_eng_kb() -> InlineKeyboardMarkup:
    kb_en = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton(text=great, callback_data='Отлично владею')
    b2 = InlineKeyboardButton(text=good, callback_data='Хорошо владею')
    b3 = InlineKeyboardButton(text=not_good, callback_data='Удовлетворительно владею')
    b4 = InlineKeyboardButton(text=bad, callback_data='Не владею английским языком')
    b5 = InlineKeyboardButton('Назад', callback_data='back_to_uz')
    kb_en.add(b1).add(b2).add(b3).add(b4).add(b5)
    return kb_en

def get_exp_kb() -> InlineKeyboardMarkup:
    kb_exp = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton(text="Есть", callback_data='Есть')
    b2 = InlineKeyboardButton(text="Нет", callback_data='Нет')
    b3 = InlineKeyboardButton(text="Назад", callback_data='back_to_eng')
    kb_exp.add(b1, b2).add(b3)
    return kb_exp
