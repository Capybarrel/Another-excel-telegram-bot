Another excel telegram bot
========================

## English:

#### Requirements:
* **telebot** and **xlrd** library installed.
* Get your token from the father of bots in telegram, and put it on **line 7**
* If you already have the id of users with allowed access, you can specify them in advance in the list **allowed_user_id on line 9**, for all others the bot will ask for a **password** that can be changed **(line 45)**

#### Description:
My boss often asks me to find a busy vlan and tell him what kind of client and subnet it is. It can be lazy to constantly run to the computer and open Excel, especially when you're not remotely accessing a computer as intended, but digging cucumbers in the yard.

#### Using:

We have a table of legal entities, with company name, address, subnet, owner, and other information. An example of this table is attached in the photo below.
When you send him a number, he looks for it in the **column 9** (can be changed on line 101, also do not forget that the numbering is from 0), and if it finds a telegram returns the information according to the specified between the lines **120 and 126** pattern, which you can change in accordance with your table.
![telebot](https://user-images.githubusercontent.com/88328046/179794826-3d412232-8f7f-4e06-b965-bbed799592f7.png)

<small>PS: I apologize for possible grammatical errors and comments in the code in Russian. I will be glad if you point out the errors, I will try to understand and correct them. </small>

## Русский:

#### Требования:

* Установленные библиотеки **telebot** и **xlrd**
* Получить свой токен у отца ботов в телеграмм, и указать его в **линии 7**
* Если у вас уже есть id пользователей с разрешенным доступом, можно указать их заранее в списке **allowed_user_id на линии 9**, для всех остальных бот будет спрашивать **пароль** который можно поменять **(линия 45)**

#### Описание:

Мой начальник часто просит найти занятый влан и сказать что это за клиент и подсеть. Постоянно бежать за компьютер и открывать эксель бывает лень особенно когда вы на удаленном доступе не сидите за компьютером как предполагалось, а копаете огурцы во дворе.

#### Использование:

У нас есть таблица юридических лиц, с названием предприятия, адресом, подсетью, вланом и другими сведениями. Пример этой таблицы прикреплен на фото ниже.
Когда вы отправляете ему число, он ищет его в столбце 9 (поменять можно на линии 101, также не забывайте что нумерация с 0), и если находит возвращает в телеграмм сведения согласно указанному между строками 120 и 126 паттерну, который вы можете поменять в соответствии с вашей таблицей.
![telebot](https://user-images.githubusercontent.com/88328046/179794826-3d412232-8f7f-4e06-b965-bbed799592f7.png)
