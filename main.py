import os

import telebot
# Кнопки
from telebot import types
# модуль работы со временем
import datetime
# модуль для работы с базой данных
# Вспомогательные данные
from additional import TOKEN, src_folder_delete, src_files, src_files_samples, src_week_report, src_week_report2
# Недельный отчет
from Night.Report import week_report
# Категории подкатегории
from Night.start import category
# Статистика операторов
from Night.Stats import stats_operators_with_missing
from Night.coefficients import coefficients
from Day.Day import timetable_day
from Day.Tags_name import tags_name

bot = telebot.TeleBot(TOKEN)

files_name = []

def view_start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard = True)
    item1 = types.KeyboardButton('День')
    item2 = types.KeyboardButton('Ночь')
    markup.add(item1, item2)
    bot.send_message(message.chat.id, f"Здравствуйте, {message.from_user.first_name} ",
                     reply_markup=markup)

# Дневные кнопки
def day_bot(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard = True)
    item1 = types.KeyboardButton('Составить график на завтра')
    item2 = types.KeyboardButton('Составить график на сегодня')
    item3 = types.KeyboardButton('Все теги')
    item4 = types.KeyboardButton('Удалить файлы')
    item5 = types.KeyboardButton('← назад')
    markup.add(item1, item2, item3, item4,item5)
    bot.send_message(message.chat.id, "Дневные функции",
                     reply_markup=markup)
# Ночные кнопки
def night_bot(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard = True)
    item1 = types.KeyboardButton('Составить недельный отчет')
    item2 = types.KeyboardButton('Составить дневной отчет')
    item3 = types.KeyboardButton('Статистика операторов с пропущенными')
    item4 = types.KeyboardButton('Статистика операторов без пропущенных')
    item5 = types.KeyboardButton('Коэффициенты')
    item6 = types.KeyboardButton('Удалить файлы')
    item7 = types.KeyboardButton('← назад')
    markup.add(item1, item2, item3, item4, item5, item6, item7)
    bot.send_message(message.chat.id, "Ночные функции",
                     reply_markup=markup)

# Получение файлов
@bot.message_handler(content_types=['document'])
def handle_docs_photo(message):
    try:
        chat_id = message.chat.id

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        src = src_files + message.document.file_name
        files_name.append(message.document.file_name)
        with open(src, 'wb') as new_file:
            new_file.write(downloaded_file)
        bot.reply_to(message, "Сохранен")
        # route(message)
    except Exception as e:
        bot.reply_to(message, e)

# Недельный отчет
def report_week(message):
    if len(files_name) == 3:
        date_now = datetime.datetime.now() - datetime.timedelta(1)
        date_7 = date_now - datetime.timedelta(6)
        start = date_7.strftime("%Y-%m-%d")
        end = date_now.strftime("%Y-%m-%d")
        week_report(files_name[0], files_name[1], files_name[2], start, end)
        with open(src_files+files_name[0], 'rb') as f1:
            bot.send_document(message.chat.id, f1)
        with open(src_week_report, 'rb') as f2:
            bot.send_document(message.chat.id, f2)
        with open(src_week_report2, 'rb') as f3:
            bot.send_document(message.chat.id, f3)
        delete_files_in_folder()
        files_name.clear()
    else:
        bot.send_message(message.chat.id, 'Добавьте 3 файла последовательно: \n Завершенные звонки операторов '
                                           'с внесенным Лист2 \n  Файл с производительностью \n'
                                           'Файл с пропущенными')

# Отчет для ночников за день
def report_day(message):
    if len(files_name) == 2:
        res = category(files_name[0], files_name[1])
        with open(src_files+files_name[0], 'rb') as f1:
            bot.send_document(message.chat.id, f1)
            bot.send_message(message.chat.id, 'Количество обработанных звонков: ' +str(res))
        delete_files_in_folder()
        files_name.clear()
    else:
        bot.send_message(message.chat.id, 'Добавьте 2 файла последовательно: \n Файл куда добавляются '
                                           'статистика по категориям \n  Файл скаченный с категориями')

# Статистика по операторам
def stats_operators_with_missing_bot(message):
    if len(files_name) == 3:
        stats_operators_with_missing(files_name[0], files_name[1], files_name[2])
        with open(src_files_samples + 'Статистика операторов с пропущенными.xlsx', 'rb') as f1:
            bot.send_document(message.chat.id, f1)
        delete_files_in_folder()
        files_name.clear()
    else:
        bot.send_message(message.chat.id, 'Добавьте 3 файла последовательно: \n Производительность операторов '
                                           '\n  Файл с оценками \n'
                                           'Файл с пропущенными')
# Подсчет коэффициентов из файла
def coefficients_bot(message):
    if len(files_name) == 1:
        bot.send_message(message.chat.id, 'Рассчитываем, ожидайте' + u'\U0001F609')
        result = coefficients(files_name[0])
        print(result)
        bot.send_message(message.chat.id, result)
        delete_files_in_folder()
        files_name.clear()
    else:
        bot.send_message(message.chat.id, 'Добавьте файл с производительностью операторов в формате csv')
def stats_operators_no_missing_bot(message):
    if len(files_name) == 3:
        stats_operators_with_missing(files_name[0], files_name[1], files_name[2], False)
        with open(src_files_samples + 'Статистика операторов без пропущенных.xlsx', 'rb') as f1:
            bot.send_document(message.chat.id, f1)
        delete_files_in_folder()
        files_name.clear()
    else:
        bot.send_message(message.chat.id, 'Добавьте 3 файла последовательно: \n Производительность операторов '
                                           '\n  Файл с оценками \n'
                                           'Пустой файл с пропущенными')
def delete_files_in_folder():
    for filename in os.listdir(src_folder_delete):
        file_path = os.path.join(src_folder_delete, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f'Ошибка при удалении файла {file_path}. {e}')

# Составление графика на день
def timetable_day_bot(message):
    try:
        bot.send_message(message.chat.id, 'График на ' + str(int(datetime.datetime.now().day)+1)+'.' +
                         str(datetime.datetime.now().month) + '.' + str(datetime.datetime.now().year) + ' создается')
        print('График на ' + str(int(datetime.datetime.now().day)+1)+'.' +
                         str(datetime.datetime.now().month) + '.' + str(datetime.datetime.now().year) + ' создается')
        result = timetable_day()
        bot.send_message(message.chat.id, result)
        print(result)
    except Exception as e:
        bot.send_message(message.chat.id, f'Ошибка при создании графика . {e}')
        print(f'Ошибка при создании графика . {e}')

def timetable_day_today_bot(message):
    try:
        bot.send_message(message.chat.id, 'График на ' + str(int(datetime.datetime.now().day))+'.' +
                         str(datetime.datetime.now().month) + '.' + str(datetime.datetime.now().year) + ' создается')
        print('График на ' + str(int(datetime.datetime.now().day))+'.' +
                         str(datetime.datetime.now().month) + '.' + str(datetime.datetime.now().year) + ' создается')
        result = timetable_day(0)
        bot.send_message(message.chat.id, result)
        print(result)
    except Exception as e:
        bot.send_message(message.chat.id, f'Ошибка при создании графика . {e}')
        print(f'Ошибка при создании графика . {e}')

# Теги всех оперов
def tag_name_bot(message):
    result = tags_name()
    bot.send_message(message.chat.id, result)
@bot.message_handler(commands=['start'])
def start(message):
    view_start(message)

@bot.message_handler(func=lambda message: True)
def bot_message(message):
    if message.text == 'День':
        day_bot(message)
    elif message.text == 'Ночь':
        night_bot(message)
    if message.text == 'Составить недельный отчет':
        report_week(message)
    elif message.text == 'Составить дневной отчет':
        report_day(message)
    elif message.text == 'Удалить файлы':
        delete_files_in_folder()
        bot.reply_to(message, "Файлы удалены")
    elif message.text == 'Составить график на завтра':
        timetable_day_bot(message)
    elif message.text == 'Составить график на сегодня':
        timetable_day_today_bot(message)
    elif message.text == 'Статистика операторов с пропущенными':
        stats_operators_with_missing_bot(message)
    elif message.text == 'Статистика операторов без пропущенных':
        stats_operators_no_missing_bot(message)
    elif message.text == 'Коэффициенты':
        coefficients_bot(message)
    elif message.text == 'Все теги':
        tag_name_bot(message)
    elif message.text == '← назад':
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton('День')
        item2 = types.KeyboardButton('Ночь')
        markup.add(item1, item2)
        bot.send_message(message.chat.id, '← назад',
                         reply_markup=markup)

def main():
    bot.infinity_polling()

if __name__ == '__main__':
    try:
        main()
    # если возникла ошибка — сообщаем про исключение и продолжаем работу
    except Exception as e:
        print('❌❌❌❌❌ Сработало исключение! ❌❌❌❌❌')
        print(e)