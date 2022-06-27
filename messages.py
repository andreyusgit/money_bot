# Файл со всеми командами, если нужно добавить команду, то лучше добавить её сюда

from utils import TestStates

help_message = 'Этот бот ведет учет задолженностей среди пользователей чатов, в который он добавлен. ' \
                'Для подтверждения своего участия в учете долгов необходимо написать боту' \
                ' в личные сообщения /start, а также нажать стартовую кнопку в приветственном сообщении, ' \
                'который бот отправляет в беседу при добавлении в нее.\n\n' \
                'С помощью бота ты можешь выбирать людей из чата и устанавливать им сумму долга. ' \
                'Также тебе будут видны личные долги, которые ты еще не отдал. ' \
                'При установлении и возврате долга пользователю он получит уведомление для подтверждения ' \
                'твоего действия со своей стороны. Обрати внимание, что для уменьшения количества транзакций' \
                ' бот компенсирует денежные суммы между двумя людьми путем вычета одного из другого.'

hello_message = 'Привет! Я бот MY_MONEY - слежу за твоими должниками! И за тобой…\n\n' \
                'Чтобы подтвердить свое участие в учете долгов среди пользователей в этом чате, ' \
               'напиши мне в личные сообщения /start, \nа также нажми --> /add_me'

start_message = 'Стартуем!\n\n/help <- чтобы узнать подробнее что я умею'

state_success_message = 'Введена некорректная сумма'

dont_know_command = "К сожалению, я пока не знаю такую команду"

MESSAGES = {
    'start': start_message,
    'hello': hello_message,
    'help': help_message,
    'state_change': state_success_message,
    'no_command': dont_know_command
}

