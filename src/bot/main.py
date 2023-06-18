from handlers import Bot
from dotenv import load_dotenv  # загрузка переменных окружения из .env

if __name__ == '__main__':
    # создание экземпляра и запуск бота
    bot = Bot()

    # отправка сообщения message_text,receiver_user_id
    #bot.send_message(message_text="Exit")
    bot.download_file_topic()
    bot.pars_excel()
    bot.send_message(message_text="111")
