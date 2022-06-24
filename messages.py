# Файл со всеми командами, если нужно добавить команду, то лучше добавить её сюда

from utils import TestStates

help_message = 'Я написан чтобы помочь тебе на отборе, сначала введи пароль, потом система отпределит твой институт.' \
               'После этого можно будет осуществлять поиск людей, но начнем с пароля.' \
               'Для того, чтобы ввести пароль, ' \
               f'отправь команду "/password x", где x - пароль из цифр.' \
               ' Чтобы сбросить пароль, отправь \n"/password" без аргументов.' \

hello_message = 'Привет! \nНапиши мне в ЛС чтобы тебя добавили в систему'


start_message = 'Привет! Тебя приветствует бот для Адапетров.\nНажми -> /help <- чтобы узнать подробнее что я умею'
invalid_key_message = 'Пароль "{key}" не подходит.\n'
state_change_success_message = 'Введён корректный пароль\nТвой институт - "{key}" \nПросто введи Фамилию ' \
                               'Имя Отчество через пробелы и я попробую его найти \n' \
                               'чтобы сменить институт выполните сброс пароля /password'
state_reset_message = 'Пароль успешно сброшен'
current_state_message = 'Текущий институт - "{current_state}""'
thanks = "Спасибо, что воспользовались ботом!\nЛучшей поддержкой будет подписка\nна мой инстаграм - andrey_us_"
dont_know_command = "К сожаленю, я не пока не зеаю такую команду"

MESSAGES = {
    'start': start_message,
    'hello': hello_message,
    'help': help_message,
    'invalid_key': invalid_key_message,
    'state_change': state_change_success_message,
    'state_reset': state_reset_message,
    'current_state': current_state_message,
    'thx': thanks,
    'no_command': dont_know_command
}
