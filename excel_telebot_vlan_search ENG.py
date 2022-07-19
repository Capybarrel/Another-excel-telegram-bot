import telebot      # to interact with telegram API
import random       # Random Number Library
import re           # Regular expression library
import xlrd         # Library for working with excel files

# telebot.apihelper.proxy = {'https': 'https://x.x.x.x:x'}   # If you need a proxy, remove the grid at the beginning
bot = telebot.TeleBot('lalalala:blablablabla')    # token to be obtained from the bot father in the telegram
global allowed_user_id, row, vlan_by_user, vlan_counter
allowed_user_id = [123456789, 123456789, 123456789, 123456789, 123456789, 123456789]   # user database of users who are initially allowed access
trigger_next_vlan = 0

file = 'G\the path\to the\excel file.xls'
excel_file = xlrd.open_workbook(file)
current_list = excel_file.sheet_by_name('required sheet in excel file')


def login_notifications(message):
    # For added security, have the bot notify us of correct and incorrect authentications
    bot.send_message(213681179, f'Hi! Just now a user with id: {message.from_user.id} '
                                f'and username: @{message.from_user.username} '
                                f'entered the answer to the question: {message.text}')


@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):

    company_name_people_count = str(random.randint(500, 600))    # Random number generator
    bot.reply_to(message, f'Hi, nice to meet you {message.from_user.first_name}. '
                          f'My name is Shlyopa and I am one of {company_name_people_count} employees of the company_name. '
                          f'Before answering you, I have to make sure you are an employee, too.  '
                          f'Name a favorite dish of the companys employees company_name:')
    bot.register_next_step_handler(message, authentication)


@bot.message_handler(content_types=['text'])
def authentication(message):

    # If the user is in the database of allowed users, then proceed to the procedure of communicating with him.
    if message.from_user.id in allowed_user_id:
        bot.register_next_step_handler(message, user_communication)
        user_communication(message)

    else:   # If not, we require a password.
        company_name_people_count = str(random.randint(500, 600))
        if re.match('[Dd]umplings', message.text.lower()):     #  Our password with a capital or small letter
            bot.reply_to(message, f'Mmmmm, {company_name_people_count} the employees of the_company_name also love {message.text}. '
                                  f'I can see that you are our man, {message.from_user.first_name}, so I will help you out. '
                                  f'Give me the number of the vlan and I will tell you what it is for:')
            # If the answer is correct, add the employee to the list of allowed employees and move on to the next procedure
            allowed_user_id.append(message.from_user.id)
            bot.register_next_step_handler(message, user_communication)

        else:
            bot.reply_to(message, f'{message.from_user.first_name}, you are an impostor! none of the {company_name_people_count} '
                                  f'company_name employees would never eat {message.text} !')

        login_notifications(message)    # The function sends us a message about a new authorization attempt.


@bot.message_handler(content_types=['text'])
def user_communication(message):

    global vlan_by_user, trigger_next_vlan

    if message.text:
        if message.text.isdigit():
            if int(message.text) in range(2, 4094):
                vlan_by_user = str(message.text)
                bot.send_message(message.from_user.id, 'Searching...')
                trigger_next_vlan = 0
                read_excel_file(message)
            else:
                bot.send_message(message.from_user.id, f'You are a bad employee, {message.from_user.first_name}. '
                                                       f'Vlan cannot be less than 2 or greater than 4094')

        elif re.match('[да|yes]', message.text.lower()) and trigger_next_vlan > 0:

            trigger_next_vlan = 0
            read_excel_file(message, first_row=row)

        elif re.match('[не|no]', message.text.lower()) and trigger_next_vlan > 0:

            trigger_next_vlan = 0
            bot.send_message(message.from_user.id, 'No way, ask about the other vlan.')

        elif message.text.lower() == 'banana donut' and message.from_user.id == 123456789:

            bot.send_message(message.from_user.id, 'Secret password entered, finishing the job!')
            bot.stop_bot()

        else:
            bot.send_message(message.from_user.id, 'Arrrrr! There can not be any letters in the vlan!')


def read_excel_file(message, first_row=0):

    global trigger_next_vlan, row, vlan_counter
    vlan_counter = 0
    # Run the loop by rows range is taken by the number of filled rows, (but the first row can be set).
    for row in range(first_row, current_list.nrows):
        vlan = current_list.cell_value(rowx=row, colx=9)   # The value of the cell from the current row and column 10 in the loop.
        vlan_search = re.findall('\d+', str(vlan))              # Looking for all the numbers in the cell
        if vlan_by_user in vlan_search:                    # If one of the numbers is right
            vlan_counter += 1
            if trigger_next_vlan == 0:  # If this is the first vlan found.
                information_about_the_vlan(message, current_list, row)  # Display the message.
            else:
                bot.send_message(message.from_user.id, 'Shlyopa found another entry on this vin, shall I show it to you?')
                break

            trigger_next_vlan = 1

    if vlan_counter == 0:
        bot.send_message(message.from_user.id, 'Shlyopa did not find anything.')
        trigger_next_vlan = 0


def information_about_the_vlan(message, current_list, row):

    bot.send_message(message.from_user.id, f'\nClient Name:  {current_list.cell_value(rowx=row, colx=3)}'
                                           f'\nCustomer Address:  {current_list.cell_value(rowx=row, colx=4)} '
                                           f' {current_list.cell_value(rowx=row, colx=5)}'
                                           f'\nType of service: {current_list.cell_value(rowx=row, colx=6)}'
                                           f'\nVpls or termination point: {current_list.cell_value(rowx=row, colx=8)}'
                                           f'\nSpeed Mb/s: {current_list.cell_value(rowx=row, colx=10)}'
                                           f'\nClient network: {current_list.cell_value(rowx=row, colx=11)}')


bot.polling()
input('Press any button to close the bot.')