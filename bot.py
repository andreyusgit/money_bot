import asyncio
import logging
from contextlib import suppress

import openpyxl
from aiogram.utils.exceptions import MessageCantBeDeleted, MessageToDeleteNotFound
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
all_users = openpyxl.Workbook()
worksheet = all_users['Sheet']
worksheet['A1']='---'
all_users.save('users.xlsx')


async def delete_message(message: types.Message, sleep_time: int = 0):
    await asyncio.sleep(sleep_time)
    with suppress(MessageCantBeDeleted, MessageToDeleteNotFound):
        await message.delete()


@dp.message_handler(content_types=["new_chat_members"])
async def new_member(message: types.Message):
    await message.delete()
    await message.answer(MESSAGES['hello'])


@dp.message_handler(commands=['add_me'])
async def process_add_user_command(message: types.Message):
    if str(message.chat.type) == 'group':
        username = message.from_user.username
        name_table = str(abs(message.chat.id))
        if "/add_me" in message.text:
            await message.delete()
        wb = openpyxl.load_workbook(f'{name_table}.xlsx')
        first_sheet = wb['main']
        second_sheet = wb['agrees']
        for i in range(2, 1000):
            if first_sheet[f'A{i}'].value == username:
                msg = await bot.send_message(message.chat.id, f'@{username} ты уже есть в базе')
                asyncio.create_task(delete_message(msg, 5))
                break
            elif first_sheet[f'A{i}'].value is None:
                coll = get_column_letter(i)
                first_sheet[f'A{i}'] = username
                first_sheet[f'{coll}1'] = username
                second_sheet[f'A{i}'] = username
                second_sheet[f'{coll}1'] = username
                for j in range(2, i + 1):
                    first_sheet[f'{get_column_letter(j)}{i}'] = 0
                    first_sheet[f'{get_column_letter(i)}{j}'] = 0
                    second_sheet[f'{get_column_letter(j)}{i}'] = 0
                    second_sheet[f'{get_column_letter(i)}{j}'] = 0
                wb.save(f"{abs(int(name_table))}.xlsx")
                users = load_workbook('users.xlsx')
                sheet = users.worksheets[0]
                row_count = sheet.max_row
                sheet[f'A{row_count + 1}'] = username
                sheet[f'B{row_count + 1}'] = name_table
                users.save('users.xlsx')
                break


@dp.message_handler(commands=['start_me'])
async def process_start_command(message: types.Message):
    if str(message.chat.type) == 'group':
        name_table = str(abs(message.chat.id))
        try:
            wb = Workbook()
            ws1 = wb.create_sheet("Sheet_A")
            ws1.title = "main"
            ws1 = wb.create_sheet("Sheet_B")
            ws1.title = "agrees"
            wb.remove_sheet(wb["Sheet"])
            wb.save(f"{abs(int(name_table))}.xlsx")
        except Exception as ex:
            await bot.send_message(message.chat.id, str(ex))
        await bot.send_message(message.chat.id, MESSAGES['hello'])


@dp.message_handler(commands=['add_debts'])
async def process_add_user_command(message: types.Message):
    name_table = 'название таблицы'
    wb = openpyxl.load_workbook(f'{name_table}.xlsx')
    current_sheet = wb['main']
    coll = 1
    for i in range(2, 1000):
        if current_sheet[f'{get_column_letter(i)}1'].value is None:
            break
        if current_sheet[f'{get_column_letter(i)}1'].value == message.from_user.username:
            coll=get_column_letter(i)
            break
    for i in range(2, 1000):
        if current_sheet[f'A{i}'].value is None:
            break
        if current_sheet[f'A{i}'].value == 'юзер один из выбранных':
            current_sheet[f'{coll}{i}'] = current_sheet[f'{coll}{i}'].value + 'долг'/'количество выбранных'


@dp.message_handler(commands=['delete_debts'])
async def process_add_user_command(message: types.Message):
    name_table = 'название таблицы'
    wb = openpyxl.load_workbook(f'{name_table}.xlsx')
    current_sheet = wb['main']
    row=1
    for i in range(2, 1000):
        if current_sheet[f'A{i}'].value is None:
            break
        if current_sheet[f'A{i}'].value == message.from_user.username:
            row=i
            break
    for i in range(2, 1000):
        if current_sheet[f'{get_column_letter(i)}1'].value is None:
            break
        if current_sheet[f'{get_column_letter(i)}1'].value == 'юзер выбранный':
            current_sheet[f'{get_column_letter(i)}{row}'] = current_sheet[f'{get_column_letter(i)}{row}'].value - 'долг'
            break


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
