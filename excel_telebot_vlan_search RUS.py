import telebot      # Бот для взаимодействия с telegram API
import random       # Библиотека случайных чисел
import re           # Библиотека регулярных выражений
import xlrd         # Библиотека работы с excel файлами

# telebot.apihelper.proxy = {'https': 'https://x.x.x.x:x'}             # Если нужно прокси убери решетку в начале
bot = telebot.TeleBot('lalalala:blablablabla')     # Наш токен
global allowed_user_id, row, vlan_by_user, vlan_counter
allowed_user_id = [123456789, 123456789, 123456789, 123456789, 123456789, 123456789] # база пользователей которым изначально разрешен доступ
trigger_next_vlan = 0

file = 'G:\Патч к необходимому файлу\excel.xls'
excel_file = xlrd.open_workbook(file)
current_list = excel_file.sheet_by_name('Необходимый лист в эксель файле')


def login_notifications(message):
    # В целях дополнительной безопасности пусть бот уведомляет нас о верных и неверных аутентификациях
    bot.send_message(213681179, f'Привет! Только что пользователь c id: {message.from_user.id} '
                                f'и username: @{message.from_user.username} '
                                f'ввёл ответ на вопрос: {message.text}')


@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):

    company_name_people_count = str(random.randint(500, 600))    # Генератор случайных чисел
    bot.reply_to(message, f'Привет, приятно познакомиться {message.from_user.first_name}. '
                          f'Меня зовут Шлёпа, я один из {company_name_people_count} сотрудников имя_компании (вставь свою компанию). '
                          f'Прежде чем отвечать тебе, я должен убедится что ты тоже сотрудник. '
                          f'Назови любимое блюдо сотрудников имя_компании:')
    bot.register_next_step_handler(message, authentication)


@bot.message_handler(content_types=['text'])
def authentication(message):

    # Если пользователь есть в базе разрешенных пользователей, то переходим к процедуре общения с ним.
    if message.from_user.id in allowed_user_id:
        bot.register_next_step_handler(message, user_communication)
        user_communication(message)

    else:   # Если нет то требуем пароль.
        company_name_people_count = str(random.randint(500, 600))
        if re.match('[Пп]ельмени', message.text.lower()):     # Наш пароль с маленькой или большой буквы
            bot.reply_to(message, f'Ммммм, {company_name_people_count} сотрудников имя_компании тоже любят {message.text}. '
                                  f'Я вижу что ты наш человек, {message.from_user.first_name}, поэтому помогу тебе. '
                                  f'Назови номер влана и я расскажу зачем он нужен:')
            # В случае правильного ответа добавляем сотрудника в список разрешенных и переходим к следующей процедуре
            allowed_user_id.append(message.from_user.id)
            bot.register_next_step_handler(message, user_communication)

        else:
            bot.reply_to(message, f'{message.from_user.first_name}, ты самозванец! ни один из {company_name_people_count} '
                                  f'сотрудников имя_компании никогда бы не стал есть {message.text} !')

        login_notifications(message)    # Функция отправляет нам сообщение о новой попытке авторизации.


@bot.message_handler(content_types=['text'])
def user_communication(message):

    global vlan_by_user, trigger_next_vlan

    if message.text:
        if message.text.isdigit():
            if int(message.text) in range(2, 4094):
                vlan_by_user = str(message.text)
                bot.send_message(message.from_user.id, 'Произвожу поиск...')
                trigger_next_vlan = 0
                read_excel_file(message)
            else:
                bot.send_message(message.from_user.id, f'Ты плохой сотрудник, {message.from_user.first_name}. '
                                                       f'Влан не может быть меньше 2 или больше 4094')

        elif re.match('[да|yes]', message.text.lower()) and trigger_next_vlan > 0:

            trigger_next_vlan = 0
            read_excel_file(message, first_row=row)

        elif re.match('[не|no]', message.text.lower()) and trigger_next_vlan > 0:

            trigger_next_vlan = 0
            bot.send_message(message.from_user.id, 'На нет и суда нет, спрашивай по другому влану.')

        elif message.text.lower() == 'банановый донат' and message.from_user.id == 123456789:

            bot.send_message(message.from_user.id, 'Ввёдён секретный пароль, завершаю работу!')
            bot.stop_bot()

        else:
            bot.send_message(message.from_user.id, 'Арррр! В влане не может быть букв!')


def read_excel_file(message, first_row=0):

    global trigger_next_vlan, row, vlan_counter
    vlan_counter = 0
    # Запускаем цикл по строкам range берется по количеству заполненных строк, (но первую строку можно задать).
    for row in range(first_row, current_list.nrows):
        vlan = current_list.cell_value(rowx=row, colx=9)   # Значение ячейки из текущей по циклу строки и 10 столбца.
        vlan_search = re.findall('\d+', str(vlan))              # Ищем все числа в ячейке
        if vlan_by_user in vlan_search:                    # Если одно из чисел нужное
            vlan_counter += 1
            if trigger_next_vlan == 0:  # Если это первый найденный влан.
                information_about_the_vlan(message, current_list, row)  # Выводим сообщение.
            else:
                bot.send_message(message.from_user.id, 'Шлёпа нашёл еще одну запись по этому влану, показать?')
                break

            trigger_next_vlan = 1

    if vlan_counter == 0:
        bot.send_message(message.from_user.id, 'Шлёпа ничего не нашёл')
        trigger_next_vlan = 0


def information_about_the_vlan(message, current_list, row):

    bot.send_message(message.from_user.id, f'\nНазвание клиента:  {current_list.cell_value(rowx=row, colx=3)}'
                                           f'\nАдрес клиента:  {current_list.cell_value(rowx=row, colx=4)} '
                                           f' {current_list.cell_value(rowx=row, colx=5)}'
                                           f'\nТип сервиса: {current_list.cell_value(rowx=row, colx=6)}'
                                           f'\nVpls или точка терминации: {current_list.cell_value(rowx=row, colx=8)}'
                                           f'\nСкорость Mb/s: {current_list.cell_value(rowx=row, colx=10)}'
                                           f'\nПодсеть клиента: {current_list.cell_value(rowx=row, colx=11)}')


bot.polling()
input('Нажмите любую кнопку чтобы закрыть бота.')