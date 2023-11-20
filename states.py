from aiogram.dispatcher.filters.state import StatesGroup, State


class ProfileStatesGroup(StatesGroup):
    initial_keyboards = State()
    cause_of_rejection = State()
    input_number = State()
    input_name = State()
    input_surname = State()
    input_birthday = State()
    input_day = State()
    input_month = State()
    input_year = State()
    input_Tashkent_or_other_town = State()
    input_district = State()
    input_other_town_and_district = State()
    input_edu = State()
    input_rus = State()
    input_uzb = State()
    input_eng = State()
    input_experience = State()
    experience_describe = State()
    input_day_and_time = State()

class AdminStatesGroup(StatesGroup):
    chat_id = State()
    message = State()
