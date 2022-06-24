import logging

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

from aiogram import Bot, types
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.types import ReplyKeyboardRemove, \
    ReplyKeyboardMarkup, KeyboardButton, \
    InlineKeyboardMarkup, InlineKeyboardButton

from config import TOKEN
from utils import TestStates
from messages import MESSAGES

logging.basicConfig(format=u'%(filename)+13s [ LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
                    level=logging.DEBUG)

bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())


@dp.message_handler(content_types=["new_chat_members"])
async def new_member(message: types.Message):
    await message.delete()
    await message.answer(MESSAGES['hello'])


@dp.message_handler(commands=['add_me'])
async def process_add_user_command(message: types.Message):
    name_table = str(abs(message.chat.id))
    if "/add_me" in message.text:
        await message.delete()
    file = openpyxl.load_workbook(f'{name_table}.xlsx')
    current_sheet = file.get_sheet_by_name('main')
    Column_C = current_sheet['C']
    try:
        max_row_for_c = len(Column_C) + 1
    except ValueError:
        max_row_for_c = 2
    coll = get_column_letter(max_row_for_c)
    print(max_row_for_c)
    print(coll)

    current_sheet[f'A{max_row_for_c}'] = message.from_user.username
    current_sheet[f'{coll}1'] = message.from_user.username
    file.save(f"{abs(int(name_table))}.xlsx")


@dp.message_handler(commands=['start_me'])
async def process_start_command(message: types.Message):
    inline_btn_1 = InlineKeyboardButton('Хочу участвовать', callback_data='add_to_bd')
    inline_kb1 = InlineKeyboardMarkup().add(inline_btn_1)
    name_table = str(abs(message.chat.id))
    try:
        wb = Workbook()
        ws1 = wb.create_sheet("Sheet_A")
        ws1.title = "main"
        ws1 = wb.create_sheet("Sheet_B")
        ws1.title = "agrees"
        wb.remove_sheet(wb.get_sheet_by_name("Sheet"))
        wb.save(f"{abs(int(name_table))}.xlsx")
    except Exception as ex:
        await bot.send_message(message.chat.id, str(ex))
    await bot.send_message(message.chat.id, "Ты открыл ящик пандоры, хуйлуша\nЧтобы записать тебя в список пидарасов "
                                            "нажми --> /add_me\nТак же, чтобы тебя, пидрилу ушастую, "
                                            "можно было найти и предъявить за дол тебе, голодранец, нужно написать "
                                            "мне в личные сообщения /star"
                                            "\nЗапомнил, петушара ?", reply_markup=inline_kb1)


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
