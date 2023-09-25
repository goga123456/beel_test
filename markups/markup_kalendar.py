from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

def get_birthday_kb() -> InlineKeyboardMarkup:
    kbirth = InlineKeyboardMarkup(resize_keyboard=True)
    b1 = InlineKeyboardButton('День', callback_data='day')
    b2 = InlineKeyboardButton('Месяц', callback_data='month')
    b3 = InlineKeyboardButton('Год', callback_data='year')
    b4 = InlineKeyboardButton('Назад', callback_data='back_to_surname')
    b5 = InlineKeyboardButton('Отправить', callback_data='send_birth')
    kbirth.add(b1, b2, b3).add(b4, b5)
    return kbirth


def get_birthday_day_kb() -> InlineKeyboardMarkup:
    markup_calendar_day = InlineKeyboardMarkup(resize_keyboard = True, row_width=7)
    item1 = InlineKeyboardButton('1', callback_data='1')
    item2 = InlineKeyboardButton('2', callback_data='2')
    item3 = InlineKeyboardButton('3', callback_data='3')
    item4 = InlineKeyboardButton('4', callback_data='4')
    item5 = InlineKeyboardButton('5', callback_data='5')
    item6 = InlineKeyboardButton('6', callback_data='6')
    item7 = InlineKeyboardButton('7', callback_data='7')
    item8 = InlineKeyboardButton('8', callback_data='8')
    item9 = InlineKeyboardButton('9', callback_data='9')
    item10 = InlineKeyboardButton('10', callback_data='10')
    item11 = InlineKeyboardButton('11', callback_data='11')
    item12 = InlineKeyboardButton('12', callback_data='12')
    item13 = InlineKeyboardButton('13', callback_data='13')
    item14 = InlineKeyboardButton('14', callback_data='14')
    item15 = InlineKeyboardButton('15', callback_data='15')
    item16 = InlineKeyboardButton('16', callback_data='16')
    item17 = InlineKeyboardButton('17', callback_data='17')
    item18 = InlineKeyboardButton('18', callback_data='18')
    item19 = InlineKeyboardButton('19', callback_data='19')
    item20 = InlineKeyboardButton('20', callback_data='20')
    item21 = InlineKeyboardButton('21', callback_data='21')
    item22 = InlineKeyboardButton('22', callback_data='22')
    item23 = InlineKeyboardButton('23', callback_data='23')
    item24 = InlineKeyboardButton('24', callback_data='24')
    item25 = InlineKeyboardButton('25', callback_data='25')
    item26 = InlineKeyboardButton('26', callback_data='26')
    item27 = InlineKeyboardButton('27', callback_data='27')
    item28 = InlineKeyboardButton('28', callback_data='28')
    item29 = InlineKeyboardButton('29', callback_data='29')
    item30 = InlineKeyboardButton('30', callback_data='30')
    item31 = InlineKeyboardButton('31', callback_data='31')
    markup_calendar_day.add(item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12,
                            item13,
                            item14, item15, item16, item17, item18, item19, item20, item21, item22, item23, item24,
                            item25,
                            item26, item27, item28, item29, item30, item31)
    return markup_calendar_day


def get_birthday_month_kb() -> InlineKeyboardMarkup:
    markup_calendar_month = InlineKeyboardMarkup(row_width=4)
    item1 = InlineKeyboardButton('Январь', callback_data='0 1')
    item2 = InlineKeyboardButton('Февраль', callback_data='0 2')
    item3 = InlineKeyboardButton('Март', callback_data='0 3')
    item4 = InlineKeyboardButton('Апрель', callback_data='0 4')
    item5 = InlineKeyboardButton('Май', callback_data='0 5')
    item6 = InlineKeyboardButton('Июнь', callback_data='0 6')
    item7 = InlineKeyboardButton('Июль', callback_data='0 7')
    item8 = InlineKeyboardButton('Август', callback_data='0 8')
    item9 = InlineKeyboardButton('Сентябрь', callback_data='0 9')
    item10 = InlineKeyboardButton('Октябрь', callback_data='1 0')
    item11 = InlineKeyboardButton('Ноябрь', callback_data='1 1')
    item12 = InlineKeyboardButton('Декабрь', callback_data='1 2')
    markup_calendar_month.add(item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11,
                              item12)
    return markup_calendar_month


def get_birthday_year_kb() -> InlineKeyboardMarkup:
    markup_calendar_year = InlineKeyboardMarkup(row_width=5)
    item1 = InlineKeyboardButton('1970', callback_data='1970')
    item2 = InlineKeyboardButton('1971', callback_data='1971')
    item3 = InlineKeyboardButton('1972', callback_data='1972')
    item4 = InlineKeyboardButton('1973', callback_data='1973')
    item5 = InlineKeyboardButton('1974', callback_data='1974')
    item6 = InlineKeyboardButton('1975', callback_data='1975')
    item7 = InlineKeyboardButton('1976', callback_data='1976')
    item8 = InlineKeyboardButton('1977', callback_data='1977')
    item9 = InlineKeyboardButton('1978', callback_data='1978')
    item10 = InlineKeyboardButton('1979', callback_data='1979')
    item11 = InlineKeyboardButton('1980', callback_data='1980')
    item12 = InlineKeyboardButton('1981', callback_data='1981')
    item13 = InlineKeyboardButton('1982', callback_data='1982')
    item14 = InlineKeyboardButton('1983', callback_data='1983')
    item15 = InlineKeyboardButton('1984', callback_data='1984')
    item16 = InlineKeyboardButton('1985', callback_data='1985')
    item17 = InlineKeyboardButton('1986', callback_data='1986')
    item18 = InlineKeyboardButton('1987', callback_data='1987')
    item19 = InlineKeyboardButton('1988', callback_data='1988')
    item20 = InlineKeyboardButton('1989', callback_data='1989')
    item21 = InlineKeyboardButton('1990', callback_data='1990')
    item22 = InlineKeyboardButton('1991', callback_data='1991')
    item23 = InlineKeyboardButton('1992', callback_data='1992')
    item24 = InlineKeyboardButton('1993', callback_data='1993')
    item25 = InlineKeyboardButton('1994', callback_data='1994')
    item26 = InlineKeyboardButton('1995', callback_data='1995')
    item27 = InlineKeyboardButton('1996', callback_data='1996')
    item28 = InlineKeyboardButton('1997', callback_data='1997')
    item29 = InlineKeyboardButton('1998', callback_data='1998')
    item30 = InlineKeyboardButton('1999', callback_data='1999')
    item31 = InlineKeyboardButton('2000', callback_data='2000')
    item32 = InlineKeyboardButton('2001', callback_data='2001')
    item33 = InlineKeyboardButton('2002', callback_data='2002')
    item34 = InlineKeyboardButton('2003', callback_data='2003')
    item35 = InlineKeyboardButton('2004', callback_data='2004')
    item36 = InlineKeyboardButton('2005', callback_data='2005')
    item37 = InlineKeyboardButton('2006', callback_data='2006')
    item38 = InlineKeyboardButton('2007', callback_data='2007')
    item39 = InlineKeyboardButton('2008', callback_data='2008')
    item40 = InlineKeyboardButton('2009', callback_data='2009')
    markup_calendar_year.add(item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12,
                             item13,
                             item14, item15, item16, item17, item18, item19, item20, item21, item22, item23, item24,
                             item25,
                             item26, item27, item28, item29, item30, item31, item32, item33, item34, item35, item36,
                             item37,
                             item38, item39, item40)
    return markup_calendar_year


