import telebot
from telebot import types
import time

# Создаем экземпляр бота
bot = telebot.TeleBot('6559164923:AAEQQTWvX9daoQlLn7Cixrmm-R9dexPT1M8')
# Функция, обрабатывающая команду /start
@bot.message_handler(commands=["start"])
def start(m, res=False):
        # Добавляем две кнопки
        markup=types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1=types.KeyboardButton("Пожать руку")
        item2=types.KeyboardButton("Покричать")
        markup.add(item1)
        markup.add(item2)
        bot.send_message(m.chat.id, 'Нажми \nПожать руку — для получения интересного эффекта\nПокричать — для получения психологической поддержки ',  reply_markup=markup)

# Получение сообщений от юзера
@bot.message_handler(content_types=["text"])
def handle_text(message):
    # Если юзер прислал 1, выдаем ему случайный факт
    if message.text.strip() == 'Пожать руку' :
            bot.send_message(message.chat.id, "Молодец, жму руку!")
    # Если юзер прислал 2, выдаем умную мысль
    elif message.text.strip() == 'Покричать':
            bot.send_message(message.chat.id, "АААААааааАА!")
            time.sleep(1)
            bot.send_message(message.chat.id, "тсссс...")
            time.sleep(3)
            bot.send_message(message.chat.id, "Нет.. все же ААААААААААААААААА")

# Запускаем бота
bot.polling(none_stop=True, interval=0)

