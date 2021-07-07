import pandas as pd
import telebot
from telebot import types
import random

data_frame = pd.read_excel('./tablee.xlsx', engine="openpyxl")
data = data_frame.to_dict('index')
data_list_tmp = []
for i in range(0, len(data)):
    data_list_tmp.append(data[i])
data_list = []
for i in range(0, len(data)):

    data_list.append(
            {
                "Название": data_list_tmp[i]['Название'],
                "Количество раундов": data_list_tmp[i]['Количество раундов'],
                "Список раундов": data_list_tmp[i]['Список раундов'].strip().split('; '),
                "Список тем": data_list_tmp[i]['Список тем'].strip().split('; '),
                "Ссылка для скачивания": data_list_tmp[i]['Ссылка для скачивания']
            }
        )


bot = telebot.TeleBot('1875185462:AAE3fB4u58Omv7ld3nWUbs6ndjdTpHa5SlQ')

MARKUP = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
ITEM1 = types.KeyboardButton('📄 Получить таблицу')
ITEM2 = types.KeyboardButton('❓ Рандомный пак')
ITEM3 = types.KeyboardButton('🔎 Поиск пака')
ITEM4 = types.KeyboardButton('❗ Помощь')
MARKUP.add(ITEM1, ITEM2, ITEM3, ITEM4)


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет, {0.first_name}! Я с удовольствием помогу тебе найти пакеты для Своей "
                                      "Игры.".format(message.from_user), reply_markup=MARKUP)


@bot.callback_query_handler(func=lambda call: True)
def ans(call):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    item1 = types.KeyboardButton('Вернуться к поиску')
    item2 = types.KeyboardButton('Главное меню')
    markup.add(item1, item2)
    if 0 < int(call.data) < len(data_list):
        bot.send_message(call.message.chat.id, "Ты решил узнать список тем для пака с названием " +
                         data_list[int(call.data)]['Название'])
        string = '📌 Список тем: \n'
        for j in range(0, len(data_list[int(call.data)]['Список тем'])):
            string = string + " " + data_list[int(call.data)]['Список тем'][j] + '\n'
        msg = bot.send_message(call.message.chat.id, string, reply_markup=markup)
        bot.register_next_step_handler(msg, answer2)


@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text.lower() == "привет":
        bot.send_message(message.chat.id, "Привет, чем я могу тебе помочь?")
    elif message.text == "📄 Получить таблицу":
        bot.send_message(message.chat.id, "Удачной игры! Если тебе лень искать пак самому, то ты можешь воспользоваться"
                                          " другими моими функциями!")
        file = open('tablee.xlsx', 'rb')
        bot.send_document(message.chat.id, file)
    elif message.text == '❓ Рандомный пак':
        k = random.randint(0, len(data_list))
        a = data_list[k]
        string = "▪Название пакета: " + a['Название'] + '\n' + "▫Ссылка для скачивания: " + a["Ссылка для скачивания"] + \
                 '\n' + "▪Количество раундов: " + str(a['Количество раундов']) + '\n' + "▫Список раундов:" + "\n"
        for i in range(0, a['Количество раундов']):
            string = string + " " + a['Список раундов'][i] + '\n'
        bot.send_message(message.chat.id, string)
        s = '📌 Список тем: \n'
        for j in range(0, len(a["Список тем"])):
            s = s + " " + a['Список тем'][j] + '\n'
        bot.send_message(message.chat.id, s)
    elif message.text == "🔎 Поиск пака":
        rmk = types.ReplyKeyboardMarkup(resize_keyboard=True)
        rmk.add(types.KeyboardButton('🔙 Назад'))
        msg = bot.send_message(message.chat.id, "Введите слово для поиска пака", reply_markup=rmk)
        bot.register_next_step_handler(msg, search)
    elif message.text == "🔙 Назад":
        bot.send_message(message.chat.id, "🔙 Назад", reply_markup=MARKUP)
    elif message.text == "❗ Помощь" or message.text == "/help":
        bot.send_message(message.chat.id, """🤖 Я - бот, который поможет тебе с поиском пакетов с вопросами для Своей Игры.
Если ты выберешь:
   📄 Получить таблицу - я отправю тебе все паки, которые у меня есть в виде таблицы Excel
   ❓ Рандомный пак - я отправлю тебе пакет с вопросами на свой вкус
   🔎 Поиск пака - мы с тобой сможем найти пакет с помощью ключевого слова, которое ты мне отправишь

Выбери, что ты хочешь сделать, в меню:)""", reply_markup=MARKUP)
    else:
        bot.send_message(message.chat.id, "Я тебя не понимаю. Напиши /help.")


def search(message):
    if message.text == "🔙 Назад":
        bot.send_message(message.chat.id, "🔙 Назад", reply_markup=MARKUP)
    else:
        if len(message.text) < 3:
            msg = bot.send_message(message.chat.id, "Для поиска введите 3 или более символа")
            bot.register_next_step_handler(msg, search)
        else:
            list_of_pac = []
            for i in range(0, len(data_list)):
                words = ''
                for j in range(0, int(data_list[i]["Количество раундов"])):
                    words += data_list[i]['Список раундов'][j] + ' '
                if message.text.lower() in data_list[i]["Название"].lower():
                    list_of_pac.append({'Пак': data_list[i], 'Номер в списке': i})
                elif message.text.lower() in words.lower():
                    list_of_pac.append({'Пак': data_list[i], 'Номер в списке': i})
            if len(list_of_pac) == 0:
                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
                item1 = types.KeyboardButton('Продолжить поиск')
                item2 = types.KeyboardButton('В главное меню')
                markup.add(item1, item2)
                msg = bot.send_message(message.chat.id,
                                       "К сожалению, паки с таким словом не были найдены. Можешь попробовать "
                                       "еще раз!", reply_markup=markup)
                bot.register_next_step_handler(msg, answer)
            else:
                for i in range(0, len(list_of_pac)):
                    string = "▪Название пакета: " + list_of_pac[i]['Пак'][
                        'Название'] + '\n' + "▫Ссылка для скачивания: " + \
                             list_of_pac[i]['Пак']["Ссылка для скачивания"] + '\n' + "▪Количество раундов: " + \
                             str(list_of_pac[i]['Пак']['Количество раундов']) + '\n' + "▫Список раундов:" + "\n"
                    for j in range(0, list_of_pac[i]['Пак']['Количество раундов']):
                        string = string + " " + list_of_pac[i]['Пак']['Список раундов'][j] + '\n'
                    markup_inline = types.InlineKeyboardMarkup()
                    item1 = types.InlineKeyboardButton(text="Узнать темы",
                                                       callback_data=str(list_of_pac[i]['Номер в списке']))
                    markup_inline.add(item1)
                    bot.send_message(message.chat.id, string, reply_markup=markup_inline)
                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
                item1 = types.KeyboardButton('Продолжить поиск')
                item2 = types.KeyboardButton('В главное меню')
                markup.add(item1, item2)
                msg = bot.send_message(message.chat.id, "⏰ Поиск окончен", reply_markup=markup)
                bot.register_next_step_handler(msg, answer)


def answer(message):
    if message.text == "В главное меню":
        bot.send_message(message.chat.id, "Возвращаюсь в главное меню", reply_markup=MARKUP)
    elif message.text == 'Продолжить поиск':
        rmk = types.ReplyKeyboardMarkup(resize_keyboard=True)
        rmk.add(types.KeyboardButton('🔙 Назад'))
        msg = bot.send_message(message.chat.id, "Введите слово для поиска пака", reply_markup=rmk)
        bot.register_next_step_handler(msg, search)


def answer2(message):
    if message.text == "Главное меню":
        bot.send_message(message.chat.id, "Возвращаюсь в главное меню", reply_markup=MARKUP)
    elif message.text == 'Вернуться к поиску':
        rmk = types.ReplyKeyboardMarkup(resize_keyboard=True)
        rmk.add(types.KeyboardButton('🔙 Назад'))
        msg = bot.send_message(message.chat.id, "Введите слово для поиска пака", reply_markup=rmk)
        bot.register_next_step_handler(msg, search)


bot.polling(none_stop=True, interval=0)
