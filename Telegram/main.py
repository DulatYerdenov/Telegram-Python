import asyncio
import logging
import sqlite3
import openai

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import BotCommand, BotCommandScopeDefault, ReplyKeyboardMarkup, KeyboardButton
from aiogram.utils.markdown import hbold

from dotenv import load_dotenv
import os

load_dotenv()


# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Токен вашего бота
TOKEN_BOT = os.getenv('TOKEN_BOT')


# ID администратора
ADMIN_ID = os.getenv('ADMIN_ID')

# Инициализация бота и диспетчера
bot = Bot(token=TOKEN_BOT)
dp = Dispatcher()

# Подключение к базе данных SQLite
conn = sqlite3.connect('my_database.db')
cursor = conn.cursor()

# Создание таблицы users, если она не существует
cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY,
        user_id INTEGER UNIQUE,
        first_name TEXT,
        last_name TEXT,
        username TEXT
    )
''')
conn.commit()

# Настройка OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')
# Команды бота
async def set_commands(bot: Bot):
    commands = [
        BotCommand(command="start", description="Начало работы"),
        BotCommand(command="help", description="Помощь"),
        BotCommand(command="lesson", description="Получить урок"),
        BotCommand(command="quiz", description="Пройти тест"),
        BotCommand(command="project", description="Идеи для проектов"),
        BotCommand(command="export", description="Выгрузить данные"),
        BotCommand(command="ask", description="Задать вопрос по Python"),
    ]
    await bot.set_my_commands(commands, scope=BotCommandScopeDefault())

# Создаем клавиатуру
buttons = [
    KeyboardButton(text='/start'),
    KeyboardButton(text='/help'),
    KeyboardButton(text='/lesson'),
    KeyboardButton(text='/quiz'),
    KeyboardButton(text='/project'),
    KeyboardButton(text='/export'),
    KeyboardButton(text='/ask'),
]
kb = ReplyKeyboardMarkup(keyboard=[buttons], resize_keyboard=True)

# Обработчик команды /start
@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    user_id = message.from_user.id
    cursor.execute("SELECT * FROM users WHERE user_id=?", (user_id,))
    user = cursor.fetchone()

    if user is None:
        cursor.execute("INSERT INTO users (user_id, first_name, last_name, username) VALUES (?, ?, ?, ?)",
                       (user_id, message.from_user.first_name, message.from_user.last_name, message.from_user.username))
        conn.commit()
        await message.answer(f"Приветствую, {hbold(message.from_user.first_name)}! \n"
                             "Вы успешно зарегистрированы!", reply_markup=kb)
    else:
        await message.answer(f"Приветствую, {hbold(message.from_user.first_name)}! \n"
                             "Рад видеть вас снова!", reply_markup=kb)

    await message.answer(
        "Я — бот для изучения Python. Вот что я умею:\n"
        "/start - Начать работу\n"
        "/help - Получить помощь\n"
        "/lesson - Получить урок\n"
        "/quiz - Пройти тест\n"
        "/project - Идеи для проектов\n"
        "/ask - Задать вопрос по Python"
    )

# Обработчик команды /help
@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    await message.answer(
        "Список доступных команд:\n"
        "/start - Начать работу\n"
        "/help - Получить помощь\n"
        "/lesson - Получить урок\n"
        "/quiz - Пройти тест\n"
        "/project - Идеи для проектов\n" 
        "/export - Выгрузить данные (только для администратора)\n"
        "/ask - Задать вопрос по Python"
    )



# Обработчик команды /lesson
@dp.message(Command("lesson"))
async def cmd_lesson(message: types.Message):
    await message.answer("Вот ваш урок: [https://metanit.com/python/tutorial/]")

# Обработчик команды /quiz
@dp.message(Command("quiz"))
async def cmd_quiz(message: types.Message):
    await message.answer("Вот ваш тест: [https://docs.google.com/document/d/1Y1YU1lFDPkYeXshVAESStlRW5AvXb_VpMDRGWSXDR0o/edit?usp=sharing]")

# Обработчик команды /project
@dp.message(Command("project"))
async def cmd_project(message: types.Message):
    projects = [
        "1. Калькулятор",
        "2. To-Do приложение",
        "3. Чат-бот",
        "4. Парсер веб-страниц",
        "5. Игра 'Угадай число'",
        "6. Конвертер валют",
        "7. Погодное приложение",
        "8. Генератор паролей"
    ]
    await message.answer("Идеи для проектов:\n" + "\n".join(projects))

# Обработчик команды /export (только для администратора)
@dp.message(Command("export"))
async def cmd_export(message: types.Message):
    if message.from_user.id == ADMIN_ID:  # Проверяем ID пользователя
        try:
            cursor.execute("SELECT * FROM users")
            users = cursor.fetchall()

            if users:
                with open("users_data.txt", "w") as file:
                    for user in users:
                        file.write(f"ID: {user[0]}, User ID: {user[1]}, Имя: {user[2]}, Фамилия: {user[3]}, Username: {user[4]}\n")
                await message.answer_document(types.FSInputFile("users_data.txt"))
            else:
                await message.answer("В базе данных нет пользователей.")

        except Exception as e:
            logger.error(f"Ошибка при выгрузке данных: {e}")
            await message.answer("Произошла ошибка при выгрузке данных.")
    else:
        await message.answer("У вас нет прав администратора для выполнения этой команды.")

# Обработчик команды /ask (вопросы к OpenAI)
@dp.message(Command("ask"))
async def cmd_ask(message: types.Message):
    await message.answer("Напишите ваш вопрос по Python, и я постараюсь на него ответить.")

# Обработчик текстовых сообщений (для вопросов к OpenAI)
@dp.message()
async def handle_message(message: types.Message):
    if message.text.startswith("/"):
        return  # Игнорируем команды

    # Отправляем вопрос в OpenAI
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  # Используем актуальную модель
            messages=[
                {"role": "system", "content": "Ты помощник, который отвечает на вопросы по Python."},
                {"role": "user", "content": f"Ответь на вопрос по Python: {message.text}"}
            ],
            max_tokens=500,  # Максимальное количество токенов в ответе
            temperature=0.7,  # Уровень креативности
        )
        answer = response.choices[0].message['content'].strip()
        await message.answer(answer)
    except Exception as e:
        logger.error(f"Ошибка при обращении к OpenAI: {e}")
        await message.answer("Произошла ошибка при обработке вашего вопроса.")

# Запуск бота
async def main():
    await set_commands(bot)
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Бот остановлен.")
    finally:
        conn.close()  # Закрываем соединение с базой данных при завершении работы