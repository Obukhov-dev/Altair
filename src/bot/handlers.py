import vk_api  # Подключаем библиотеку Vk_Api
import os  # работа с файловой системой
import requests
import openpyxl

from vk_api.utils import get_random_id  # Уникальный id предотвращение повторной отправки одинакового сообщения.
from dotenv import load_dotenv  # загрузка переменных окружения из .env


class Bot:
    # Статус текущий сессии ВК
    vk_session = None
    # Статус доступа к APIvk
    vk_api_access = None
    # Статус авторизации
    authorized = False
    # id страницы для взаимодействия
    main_user_id = None
    # Группа для Parsing а
    group_id = None
    # Обсуждение в этой группе для Parsing а
    topic_id = None

    def __init__(self):
        """Функция инициализации через доступ APIvk"""

        # Импорт переменных окружения (ключей и тп.)
        load_dotenv()

        # Авторизация
        self.vk_api_access = self.do_auth()

        if self.vk_api_access is not None:
            self.authorized = True

        # Импорт id страницы, группы, обсуждения
        self.main_user_id = os.getenv("USER_ID")
        self.group_id = os.getenv("GROUP_ID")
        self.topic_id = os.getenv("TOPIC_ID")

    def do_auth(self):
        """
        функция Authorizations бота-пользователя.

        :return:
        """

        token = os.getenv("ACCESS_TOKEN")
        token_group = os.getenv("ACCESS_TOKEN_GROUP")  # с токена группы

        try:
            self.vk_session = vk_api.VkApi(token=token)  # от кого авторизуемся
            return self.vk_session.get_api()
        except Exception as error:
            print(error)
            return None

    def send_message(self, receiver_user_id: str = None, message_text="тестовое сообщение"):
        """
        Функция отправки сообщения.

        :param receiver_user_id: id получателя сообщения (при отсутствии подставится значение из окружения переменных)
        :param message_text: текст отправляемого сообщения
        :return:
        """

        #
        if not self.authorized:
            print("CriticalError; check ACCESS_TOKEN?")
            return

        # нет id - значение из окружения переменных
        if receiver_user_id is None:
            receiver_user_id = self.main_user_id

        try:
            self.vk_api_access.messages.send(user_id=receiver_user_id, message=message_text, random_id=get_random_id())
            print(f"Сообщение для ID {receiver_user_id} текст: {message_text}")
        except Exception as error:
            print(error)

    def download_file_topic(self):
        """
        Функция для получения файла в topic
        :return:
        """

        # Проверяем, авторизован ли пользователь
        if not self.authorized:
            print("CriticalError; check ACCESS_TOKEN?")
            return

        try:
            # Получаем прямую ссылку на файл из topic a
            requestComments = self.vk_api_access.board.getComments(group_id=self.group_id, topic_id=self.topic_id)
            print(requestComments)
            getFile = (requestComments['items'][-1]['attachments'][0]['doc']['url'])
            print(getFile)

            # Проверяем, что ответ содержит данные и что последний комментарий в теме не пустой
            if 'items' not in requestComments or not requestComments['items']:
                raise ValueError('No comments found in the topic')

            # скачиваем файл
            filename = 'rasp.xlsx'
            dir_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../rasp')
            filepath = os.path.join(dir_path, filename)
            prev_filepath = os.path.join(dir_path, 'prev_' + filename)

            # создаем директорию, если она не существует
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

            # Отправляем запрос и сохраняем содержимое в файл.
            # На выходе старая и новая версия с разными именами.
            response = requests.get(getFile)
            if os.path.exists(prev_filepath):
                os.remove(prev_filepath)

            if os.path.exists(filepath):
                os.rename(filepath, prev_filepath)

            with open(filepath, 'wb') as f:
                f.write(response.content)

        except Exception as error:
            print(error)

    def pars_excel(self):
        """
        # from openpyxl import load_workbook
        :return:
        """
        book = openpyxl.open(filename='../../rasp/rasp.xlsx', read_only=True)
        sheet = book.active

        # input
        group_search = ""
        time = "время"
        classroom = "Каб."
        day_data = ""
        numbLessons = 1  # 1-9 пн, сб. Будни: 1-7

        buffer = []
        counter = 0

        # [Строка], [Столбец]
        for numb in range(1, 20):
            if sheet[numb][0].value == "Группа":
                # Получили номер строки с заголовком группа (в numb)
                for numb1 in range(1, 70):
                    # Среди строки найдем столбец с нужной группой
                    if sheet[numb][numb1].value == group_search:
                        # Стянем столбец под нужной группой
                        for numb2 in range(numb + 3, numb + 21):
                            a = sheet[numb2][numb1].value

                            if a is None:
                                buffer += "no".split()
                                continue

                            buffer += a.split("}")

                            print(buffer)

        print(buffer)