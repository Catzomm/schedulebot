from ut.modals import *
from ut.main import schedule_handler
import telebot
from telebot import types
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from ut.schedule import schedule, schedule2
from datetime import datetime, timedelta
import time
from telebot.types import InputMediaPhoto

schedule_handler.main_schedule()
schedule_handler.teacher_schedule()
t_id = read_teachers('file/teachers_telegram.xlsx')
c_id = read_teachers('file/class_telegram.xlsx')
days = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота']
days_ru = ['понедельник', 'вторник', 'среду', 'четверг', 'пятницу', 'субботу']


# Проверка на изменение в файлах
class MyHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_modified = datetime.now() - timedelta(seconds=35)
        self.last_modified2 = datetime.now() - timedelta(seconds=35)
        self.event = ''

    def on_modified(self, event):
        global list_of_teachers, list_of_class
        day = event.src_path[2:-5].strip().lower()
        if day in days:
            if datetime.now() - self.last_modified < timedelta(seconds=35) and self.event == event:
                return
            else:
                self.last_modified = datetime.now()
                self.event = event
                bot.send_message(5702272905, 'Новое расписание')

                date = datetime.now()
                new_day = timedelta((7 - date.weekday() + days.index(day)) % 7) + date
                new_date = '.'.join(str(new_day.date()).split('-')[::-1][:-1])
                schedule_handler.main(day.capitalize())
                info = [schedule_handler.new_teachers, schedule_handler.old_teachers, schedule_handler.old_of_lessons, schedule_handler.new_of_lessons]
                mailing_telegram(info, days_ru[days.index(day)].capitalize(), new_date)

        elif day == 'расписание':
            if datetime.now() - self.last_modified2 < timedelta(seconds=35) and self.event == event:
                return
            else:
                self.event = event
                self.last_modified2 = datetime.now()
                time.sleep(0.1)
                schedule_handler.main_schedule()
                list_of_class = list_class()
                list_of_teachers = list_teachers()


#Telegram bot
count_time = -1
count_time2 = -1
list_of_teachers = list_teachers()
list_of_class = list_class()
list_schedule = {}
list_time = {}

bot = telebot.TeleBot("5030673072:AAFSwQrbCmUIf8-aAjQG9zWFYNTkYD2NJVo")
user = ''


@bot.message_handler(commands=['start'])
def start(message):
    if str(message.chat.id) not in schedule_handler.teacher_id and str(message.chat.id) not in schedule_handler.class_id:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
        item3 = types.KeyboardButton("Инструкция")
        markup.add(item3)
        bot.send_message(message.chat.id, 'Вас приветствует Telegram bot \n"КФМЛ Расписание"', parse_mode='html', reply_markup=markup)

    bot.send_message(message.chat.id, 'Пройдите регистрацию', parse_mode='html')
    markup = types.InlineKeyboardMarkup(row_width=2)
    item1 = types.InlineKeyboardButton('Учитель', callback_data='Учитель')
    item2 = types.InlineKeyboardButton('Ученик', callback_data='Ученик')
    markup.add(item1, item2)
    bot.send_message(message.chat.id, 'Укажите, Вы учитель или ученик', parse_mode='html', reply_markup=markup)


@bot.message_handler(commands=['mailing'])
def getting_the_command(message):
    if message.chat.id == 1650373681:
        bot.send_message(message.chat.id, '5')
    elif message.chat.id == 5702272905:
        send = bot.send_message(message.chat.id, 'Для кого?')
        bot.register_next_step_handler(send, sending_message)


def sending_message(message):
    global user
    status = ''.join((message.text.strip() + ' ').split(' '))
    user = status
    if status == 'all' or status in list_of_class or status in list_of_teachers:
        send = bot.send_message(message.chat.id, 'Введите сообщение')
        bot.register_next_step_handler(send, sending_message_2)
    else:
        bot.send_message(message.chat.id, 'Некорректный аргумент')
        getting_the_command(message)


def sending_message_2(message):
    teachers = read_teachers('file/teachers_telegram.xlsx') | read_teachers('file/class_telegram.xlsx')
    if user == 'all':
        for i in teachers:
            for id in teachers[i]:
                try:
                    bot.send_message(id, message.text)
                except Exception as e:
                    pass
    else:
        for id in teachers[user]:
            try:
                bot.send_message(id, message.text)
            except Exception as e:
                pass


@bot.message_handler(commands=['get'])
def get(message):
    if message.chat.id == 5702272905:
        try:
            k = 0
            mess = ''
            a = t_id | c_id
            for i in a:
                mess += f'{i}: {len(a[i])} \n\n'
                k += len(a[i])
            bot.send_message(message.chat.id, mess + f'Итого: {k}', parse_mode='html')
        except Exception as e:
            pass


@bot.message_handler()
def get_user_text(message):
    if message.text == 'Повторная регистрация':
        start(message)
    elif message.text == "Инструкция":
        manual(message)
    elif message.text == "Расписание звонков":
        call_schedule(message)
    elif message.text == 'Моё расписание':
        my_schedule(message)


def manual(message):
    markup = types.InlineKeyboardMarkup(row_width=2)
    item = types.InlineKeyboardButton("Регистрация", callback_data='help_1')
    markup.add(item)
    item = types.InlineKeyboardButton("Изменения в расписании", callback_data='help_2')
    item2 = types.InlineKeyboardButton("Моё расписание", callback_data='help_3')
    item3 = types.InlineKeyboardButton("Расписание звонков", callback_data='help_4')
    item4 = types.InlineKeyboardButton(text='Тех. поддержка', url='https://t.me/kpml_schedule_support')
    markup.add(item, item2, item3, item4)
    bot.send_message(message.chat.id, help_message, parse_mode='html', reply_markup=markup)


def call_schedule(message):
    bot.send_message(message.chat.id, '<b>Расписание звонков</b>\n\n' + '\n'.join(schedule_handler.time) +
                     '\n\n<b>Расписание звонков на Понедельник</b>\n\n' + '\n'.join(schedule_handler.time0) +
                     '\n\n<b>Расписание звонков на Субботу</b>\n\n' + '\n'.join(schedule_handler.time2),
                     parse_mode='html')


def my_schedule(message):
    global count_time, count_time2
    loading = bot.send_message(message.chat.id, 'Идёт загрузка... ⌛️', parse_mode='html')
    number = teacher_id(message.chat.id)
    if number != None:
        while True:
            if True:
                count_time += 1
                time.sleep(count_time * 5)
            if number != None:
                schedule(int(number))
                pdf_file = open('file/img/Расписание.png', 'rb')
                bot.send_photo(message.chat.id, pdf_file)
                bot.delete_message(message.chat.id, loading.message_id)
                pdf_file.close()
                count_time -= 1
                break
    else:
        while True:
            number = class_id(message.chat.id)
            if True:
                count_time2 += 1
                time.sleep(count_time2 * 5)
            if number != None:
                schedule2(int(number))
                pdf_file = open('file/img/Расписание уроков.png', 'rb')
                bot.send_photo(message.chat.id, pdf_file)
                bot.delete_message(message.chat.id, loading.message_id)
                pdf_file.close()
                count_time2 -= 1
                break
            else:
                break


def mailing_telegram(info, day, date):
    old_lesson = info[2]
    new_lesson = info[3]
    new_teacher = info[0]
    old_teacher = info[1]
    new_teachers = new_teacher | new_lesson
    old_teachers = old_teacher | old_lesson
    users = read_teachers('file/teachers_telegram.xlsx') | read_teachers('file/class_telegram.xlsx')

    for teacher in users:
        mess = ''
        new_lessson = {}
        old_lessson = {}
        if not(len(users[teacher]) == 1 and (users[teacher][0] == 'None' or users[teacher][0] == '')):
            if teacher in new_teachers:
                new_lessson = sending_new(teacher, new_teachers)

            if teacher in old_teachers:
                old_lessson = sending_old(teacher, old_teachers)

            if teacher in old_teachers or teacher in new_teachers:
                mess += f'<b>Изменения в расписании на \n{day} ({date})</b>\n\n'
                mess += beautiful_schedule(old_lessson, new_lessson)
            else:
                mess = f'У Вас нет изменений в расписании на {day} ({date})'
            for id_teacher in users[teacher]:
                if id_teacher != 'None' and id_teacher != '':
                    try:
                        send_schedule(id_teacher, mess, day, date, teacher)
                    except Exception as e:
                        pass


def send_schedule(id_teacher, mess, day, date, teacher):
    global list_schedule, list_time
    if schedule_handler.new_time:
        mess_time = f'<b>Расписание звонков на </b>\n<b>{day} ({date})</b>\n\n' + '\n'.join(schedule_handler.new_time)
        if id_teacher not in list_time or (
                id_teacher in list_time and mess_time != list_time[id_teacher]):
            bot.send_message(id_teacher, mess_time, parse_mode='html')

            schedule_img(id_teacher, mess, teacher)

            list_time[id_teacher] = mess_time
            list_schedule[id_teacher] = mess
    else:
        list_time[id_teacher] = ''
    if id_teacher not in list_schedule or (
            id_teacher in list_schedule and mess != list_schedule[id_teacher]):
        if schedule_handler.new_time:
            bot.send_message(id_teacher, f'<b>Расписание звонков на </b>\n<b>{day} ({date})</b>\n\n' + '\n'.join(
                schedule_handler.new_time), parse_mode='html')

        schedule_img(id_teacher, mess, teacher)

        list_schedule[id_teacher] = mess


def schedule_img(id_teacher, mess, teacher):
    if str(id_teacher) in schedule_handler.teacher_id:
        media_1 = InputMediaPhoto(media=open(f'file/img/{schedule_handler.file}-0.png', 'rb'), caption=mess,
                                  parse_mode='html')
        media_2 = InputMediaPhoto(media=open(f'file/img/{schedule_handler.file}-1.png', 'rb'))
        media = [media_1, media_2]
        bot.send_media_group(id_teacher, media)
    else:
        if teacher[:1].isdigit:
            if int(teacher[:-1]) <= schedule_handler.last_class:
                number = 0
            else:
                number = 1
            pdf_file = open(f'file/img/{schedule_handler.file}-{number}.png', 'rb')
            bot.send_photo(id_teacher, pdf_file, caption=mess, parse_mode='html')


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    try:
        if call.data == 'help':
            markup = types.InlineKeyboardMarkup(row_width=2)
            item = types.InlineKeyboardButton("Регистрация", callback_data='help_1')
            markup.add(item)
            item = types.InlineKeyboardButton("Изменения в расписании", callback_data='help_2')
            item2 = types.InlineKeyboardButton("Моё расписание", callback_data='help_3')
            item3 = types.InlineKeyboardButton("Расписание звонков", callback_data='help_4')
            item4 = types.InlineKeyboardButton(text='Тех. поддержка', url='https://t.me/kpml_schedule_support')
            markup.add(item, item2, item3, item4)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text=help_message, parse_mode='html', reply_markup=markup)
        elif call.data == 'help_1':
            markup = types.InlineKeyboardMarkup()
            item = types.InlineKeyboardButton('< Назад', callback_data='help')
            markup.add(item)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text=help_message_1, parse_mode='html', reply_markup=markup)
        elif call.data == 'help_2':
            markup = types.InlineKeyboardMarkup()
            item = types.InlineKeyboardButton('< Назад', callback_data='help')
            markup.add(item)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text=help_message_2, parse_mode='html', reply_markup=markup)
        elif call.data == 'help_3':
            markup = types.InlineKeyboardMarkup()
            item = types.InlineKeyboardButton('< Назад', callback_data='help')
            markup.add(item)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text=help_message_3, parse_mode='html', reply_markup=markup)
        elif call.data == 'help_4':
            markup = types.InlineKeyboardMarkup()
            item = types.InlineKeyboardButton('< Назад', callback_data='help')
            markup.add(item)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text=help_message_4, parse_mode='html', reply_markup=markup)
        elif call.data == 'back':
            markup = types.InlineKeyboardMarkup(row_width=2)
            item1 = types.InlineKeyboardButton('Учитель', callback_data='Учитель')
            item2 = types.InlineKeyboardButton('Ученик', callback_data='Ученик')
            markup.add(item1, item2)
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text='Укажите, Вы учитель или ученик',
                                  reply_markup=markup)
        elif call.data == 'Учитель':
            markup = types.InlineKeyboardMarkup(row_width=1)
            for i in range(5):
                item = types.InlineKeyboardButton(list_of_teachers[i], callback_data=list_of_teachers[i])
                markup.add(item)
            item6 = types.InlineKeyboardButton('След >>', callback_data='next 1 teacher')
            markup.add(item6)
            item7 = types.InlineKeyboardButton('< Назад', callback_data='back')
            markup.add(item7)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text='Найдите себя в списке',
                                  reply_markup=markup)

        elif call.data == 'Ученик':
            markup = types.InlineKeyboardMarkup(row_width=1)
            for i in range(5):
                item = types.InlineKeyboardButton(list_of_class[i], callback_data=list_of_class[i])
                markup.add(item)
            item6 = types.InlineKeyboardButton('След >>', callback_data='next 1 student')
            markup.add(item6)
            item7 = types.InlineKeyboardButton('< Назад', callback_data='back')
            markup.add(item7)

            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id,
                                  text='Выберите класс',
                                  reply_markup=markup)

        else:
            data = call.data.split(' ')
            if data[0] == 'previous':
                if data[2] == 'teacher':
                    list_of = list_of_teachers
                else:
                    list_of = list_of_class
                n = (int(data[1]) - 2) * 5
                markup = types.InlineKeyboardMarkup(row_width=1)
                for i in range(5):
                    item = types.InlineKeyboardButton(list_of[n + i], callback_data=list_of[n + i])
                    markup.add(item)
                if data[1] == '2':
                    item2 = types.InlineKeyboardButton('След >>', callback_data=' '.join(['next', str(int(data[1]) - 1), data[2]]))
                    markup.add(item2)
                else:
                    item2 = types.InlineKeyboardButton('<< Пред', callback_data=' '.join(['previous', str(int(data[1]) - 1), data[2]]))
                    item3 = types.InlineKeyboardButton('След >>', callback_data=' '.join(['next', str(int(data[1]) - 1), data[2]]))
                    markup.row(item2, item3)
                item7 = types.InlineKeyboardButton('< Назад', callback_data='back')
                markup.add(item7)

                bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=call.message.text,
                                      reply_markup=markup)

            elif data[0] == 'next':
                if data[2] == 'teacher':
                    list_of = list_of_teachers
                else:
                    list_of = list_of_class
                n = int(data[1]) * 5
                markup = types.InlineKeyboardMarkup(row_width=1)
                if len(list_of[n:]) >= 5:
                    k = 5
                else:
                    k = len(list_of[n:])
                for i in range(k):
                    item = types.InlineKeyboardButton(list_of[n + i], callback_data=list_of[n + i])
                    markup.add(item)
                if len(list_of[n:]) <= 5:
                    item2 = types.InlineKeyboardButton('<< Пред', callback_data=' '.join(['previous', str(int(data[1]) + 1), data[2]]))
                    markup.add(item2)
                else:
                    item2 = types.InlineKeyboardButton('<< Пред', callback_data=' '.join(['previous', str(int(data[1]) + 1), data[2]]))
                    item3 = types.InlineKeyboardButton('След >>', callback_data=' '.join(['next', str(int(data[1]) + 1), data[2]]))
                    markup.row(item2, item3)
                item7 = types.InlineKeyboardButton('< Назад', callback_data='back')
                markup.add(item7)

                bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=call.message.text,
                                      reply_markup=markup)

            else:
                if call.message.text == 'Найдите себя в списке':
                    write_telegram(call.data, int(call.message.chat.id), 'file/teachers_telegram.xlsx')
                else:
                    write_telegram(call.data, int(call.message.chat.id), 'file/class_telegram.xlsx')

                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
                item = types.KeyboardButton("Повторная регистрация")
                item2 = types.KeyboardButton("Моё расписание")
                item3 = types.KeyboardButton("Инструкция")
                item4 = types.KeyboardButton("Расписание звонков")
                markup.add(item, item2, item3, item4)
                bot.send_message(call.message.chat.id, text='Спасибо за регистрацию', reply_markup=markup)

                bot.delete_message(call.message.chat.id, call.message.message_id)
                bot.delete_message(call.message.chat.id, call.message.message_id - 1)

    except Exception as e:
        print(repr(e))


from multiprocessing import Process


def tbot():
    while True:
        try:
            bot.polling(none_stop=True)
        except Exception as e:
            time.sleep(3)
            print(e)


if __name__ == "__main__":
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path='./', recursive=False)
    observer.start()
    Process(target=tbot).start()
