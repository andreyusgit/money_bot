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
main_sheet = all_users.create_sheet("Sheet_A")
main_sheet.title = "cards"
worksheet = all_users['Sheet']
worksheet['A1'] = '---'
all_users.save('users.xlsx')
user_deb = {}


async def delete_message(message: types.Message, sleep_time: int = 0):
    await asyncio.sleep(sleep_time)
    with suppress(MessageCantBeDeleted, MessageToDeleteNotFound):
        await message.delete()


@dp.message_handler(state='*', content_types=["new_chat_members"])
async def new_member(message: types.Message):
    await message.delete()
    await message.answer(MESSAGES['hello'])


@dp.message_handler(state='*', commands=['add_me'])
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
                sheet = users['Sheet']
                row_count = sheet.max_row + 1
                sheet[f'A{row_count}'] = username
                sheet[f'B{row_count}'] = name_table
                sheet[f'C{row_count}'] = message.chat.title
                sheet[f'D{row_count}'] = message.from_user.first_name
                sheet[f'E{row_count}'] = message.from_user.last_name
                users.save('users.xlsx')
                break


@dp.message_handler(state='*', commands=['start_me'])
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


@dp.message_handler(state='*', commands=['add_debts'])
async def process_add_debts_command(message: types.Message):
    if str(message.chat.type) == 'private':
        keyboard = types.ReplyKeyboardMarkup(resize_keyboard=False)
        username = message.from_user.username
        wb2 = openpyxl.load_workbook('users.xlsx')
        active_sheet = wb2['Sheet']
        user_deb[username] = [], [], []
        for i in range(1, active_sheet.max_row + 1):
            if username == active_sheet[f'A{i}'].value:
                user_deb[username][0].append(active_sheet[f'B{i}'].value)
                user_deb[username][1].append(active_sheet[f'C{i}'].value)
        for i in range(len(user_deb[username][0])):
            keyboard.add(f'{user_deb[username][1][i]}')
        await bot.send_message(message.chat.id, 'Выбери из списка ниже в какой группе находится твой должник: ',
                               reply_markup=keyboard)
        state = dp.current_state(user=message.from_user.id)
        await state.set_state(TestStates.all()[0])


@dp.message_handler(state=TestStates.TEST_STATE_0)
async def process_add_debts_command_2(message: types.Message):
    wb2 = openpyxl.load_workbook('users.xlsx')
    active_sheet = wb2['Sheet']
    name_table = ''
    keyboard2 = types.ReplyKeyboardMarkup(resize_keyboard=False)
    for i in range(1, active_sheet.max_row + 1):
        if message.text == active_sheet[f'C{i}'].value:
            name_table = str(active_sheet[f'B{i}'].value)
    user_deb[message.from_user.username][2].append(name_table)
    wb = openpyxl.load_workbook(f'{name_table}.xlsx')
    first_sheet = wb['main']
    for i in range(2, first_sheet.max_row + 1):
        keyboard2.add(active_sheet[f'A{i}'].value)
    await bot.send_message(message.chat.id, 'Выбери из списка ниже пользователя, который тебе задолжал: ',
                           reply_markup=keyboard2)
    state = dp.current_state(user=message.from_user.id)
    await state.set_state(TestStates.all()[1])


@dp.message_handler(state=TestStates.TEST_STATE_1)
async def process_add_debts_command_2(message: types.Message):
    user_deb[message.from_user.username][2].append(message.text)
    await bot.send_message(message.chat.id, 'введи сумму долга:')
    state = dp.current_state(user=message.from_user.id)
    await state.set_state(TestStates.all()[2])


@dp.message_handler(state=TestStates.TEST_STATE_2)
async def process_add_debts_command_2(message: types.Message):
    username = user_deb[message.from_user.username][2].pop()
    name_table = user_deb[message.from_user.username][2].pop()
    wb = openpyxl.load_workbook(f'{name_table}.xlsx')
    active_sheet = wb['main']
    coll = 1
    for i in range(2, 1000):
        if active_sheet[f'{get_column_letter(i)}1'].value is None:
            break
        if active_sheet[f'{get_column_letter(i)}1'].value == message.from_user.username:
            coll = i
            break
    for i in range(2, 1000):
        if active_sheet[f'A{i}'].value is None:
            break
        if active_sheet[f'A{i}'].value == username:
            active_sheet[f'{get_column_letter(coll)}{i}'] = active_sheet[
                                                               f'{get_column_letter(coll)}{i}'].value + \
                                                            int(message.text)
            if active_sheet[f'{get_column_letter(i)}{coll}'].value > 0:
                if active_sheet[f'{get_column_letter(coll)}{i}'].value > \
                        active_sheet[f'{get_column_letter(i)}{coll}'].value:
                    active_sheet[f'{get_column_letter(coll)}{i}'] = active_sheet[
                                                                        f'{get_column_letter(coll)}{i}'].value - \
                                                                    active_sheet[
                                                                        f'{get_column_letter(i)}{coll}'].value
                else:
                    active_sheet[f'{get_column_letter(coll)}{i}'] = active_sheet[
                                                                        f'{get_column_letter(i)}{coll}'].value - \
                                                                    active_sheet[
                                                                        f'{get_column_letter(coll)}{i}'].value
        break
    wb.save(f"{name_table}.xlsx")



@dp.message_handler(state='*', commands=['delete_debts'])
async def process_delete_debts_command(message: types.Message):
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


@dp.message_handler(state='*', commands=['start'])
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


@dp.message_handler(state='*', commands=['stepa_hvatit'])
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
