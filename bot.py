import logging
import sqlite3
from openpyxl import Workbook, load_workbook

from aiogram import Bot, types
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware

from config import TOKEN
from utils import TestStates
from messages import MESSAGES

logging.basicConfig(format=u'%(filename)+13s [ LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
                    level=logging.DEBUG)

bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())
conn = sqlite3.connect('users.db')


@dp.message_handler(content_types=["new_chat_members"])
async def new_member(message: types.Message):
    cur = conn.cursor()
    await message.delete()
    await message.answer(MESSAGES['hello'])
    name_table = '_' + str(abs(message.chat.id))
    cur.execute(f"INSERT INTO {name_table} VALUES ('{str(message.from_user.username)}')")
    cur.execute(f"alter table {name_table} add column '{message.from_user.username}' 'text'")
    cur.execute(f"INSERT INTO {name_table + '_approve'} VALUES ('{str(message.from_user.username)}')")
    cur.execute(f"alter table {name_table + '_approve'} add column '{message.from_user.username}' 'text'")
    conn.commit()


@dp.message_handler(commands=['create_bd'])
async def process_start_command(message: types.Message):
    cur = conn.cursor()
    name_table = '_' + str(abs(message.chat.id))
    try:
        cur.execute(f'''CREATE TABLE {name_table} (name text)''')
        cur.execute(f'''CREATE TABLE {name_table + '_approve'} (name text)''')
        conn.commit()
    except Exception as ex:
        await bot.send_message(message.chat.id, str(ex))


@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    await message.answer(MESSAGES['start'])


@dp.message_handler(state='*', commands=['help'])
async def process_help_command(message: types.Message):
    await message.answer(MESSAGES['help'])


@dp.message_handler(state='*', commands=['thanks'])
async def process_thx_command(message: types.Message):
    await message.answer(MESSAGES['thx'])


@dp.message_handler(state=TestStates.all())
async def some_test_state_case_met(message: types.Message):
    await message.answer(MESSAGES['no_command'])


@dp.message_handler()
async def echo_message(msg: types.Message):
    await bot.send_message(msg.chat.id, "че хочешь")
    await bot.send_message(msg.from_user.id, "hello")


async def shutdown(dispatcher: Dispatcher):
    await dispatcher.storage.close()
    await dispatcher.storage.wait_closed()


if __name__ == '__main__':
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button_1 = types.KeyboardButton(text="/start")
    keyboard.add(button_1)
    executor.start_polling(dp, on_shutdown=shutdown)
