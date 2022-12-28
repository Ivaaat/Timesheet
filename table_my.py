import time
import datetime
import calendar, locale
import xlsxwriter
import os
import telebot
from telebot import types
import logging
from telebot import formatting
from edit_table import edit_timesheet, export_excel_jpeg, default_timesheet
import threading

root_path = os.path.dirname(os.path.abspath(__file__))
TOKEN = "5877308961:AAE5myH4vER7-VaROqFmc0ApExdaQK0-FiU"
DICT_DAY = {}
logger = telebot.logger
telebot.logger.setLevel(logging.DEBUG)
bot = telebot.TeleBot(TOKEN)#Токен 
bot.set_my_commands(
    commands=[
        telebot.types.BotCommand("start", "timesheet"),
    ],
)

bot.send_message(377190896,'Жми /start')


@bot.message_handler(commands='start')
def timesheet_person(message):
    id_chief = 240652259
    #chief --- "chat":{"id":240652259,"first_name":"Dmitry","last_name":"Dmitriev","username":"oDDSo","type":"private"}
    # if message.chat.id == 377190896 or message.chat.id == 240652259:
    #     name = f"{message.chat.first_name} {message.chat.username}"
    # else:
    #     name = "Иванов Иван Иваныч"
    name = f"{message.chat.first_name} {message.chat.username}"
    employee_name = formatting.mbold(name)
    bot.send_message(message.chat.id, employee_name, parse_mode='MarkdownV2')
    bot.delete_my_commands()
    return timesheet_buttons(message, name)

def timesheet_buttons(message, employee_name):
    buttons = ['По умолчанию', 'Изменить табель', 'Изменить имя']
    now = datetime.datetime.now()
    markup = types.ReplyKeyboardMarkup()
    #if message.chat.id == 377190896 or message.chat.id == 240652259:
    if  message.chat.id == 240652259:
        name_button = "Выгрузить табель"
        button_chief = types.KeyboardButton(name_button)
        markup.add(button_chief)
        if now.day not in [15, calendar.mdays[now.month]]:
            flag_string = "Еще рано выгружать табель\n"
        else:
            flag_string = ""
        name_message = f"{flag_string}Сегодня {now.day}.{now.month}.{now.year}! "
        msg = bot.send_message(message.chat.id, f"{name_message}", reply_markup = markup, parse_mode="")
        return bot.register_next_step_handler(msg, get_timesheet, buttons, employee_name)
    elif message.chat.id == 377190896:
        name_button = "Выгрузить табель"
        button_chief = types.KeyboardButton(name_button)
        markup.add(button_chief)
    else:
        name = f"{message.chat.first_name} {message.chat.username}"
        bot.send_message(message.chat.id,name + " Давай до свидания!")
        return timesheet_buttons(message, employee_name)
    for button in buttons:
        timesheet_default = types.KeyboardButton(button)
        markup.add(timesheet_default)
    msg = bot.send_message(message.chat.id, f"Создай и отправь начальнику табель!\n\n\
По умолчанию - Работа в офисе, 8 часов, выходные помечены красным\n\
Изменить табель - настройка табеля под свои нуждны\n\
Изменить ФИО - изменение дефолтного ФИО", reply_markup = markup, parse_mode="")
    bot.register_next_step_handler(msg, get_timesheet, buttons, employee_name)



def get_timesheet(message, list_buttons, employee_name, back = ""):
    now = datetime.datetime.now()
    # if now.day not in [15,num_days_month]:
    #     bot.send_message(message.chat.id, f'Еще рано отправлять табель')
    #     return timesheet_buttons(message, employee_name)
    text = message.text
    if text == list_buttons[0]:
        filename = default_timesheet (employee_name = employee_name)
        #th1 = threading.Thread(target = export_excel_jpeg, args=(filename, employee_name), name='JPG').start()
        filename_jpeg = export_excel_jpeg(filename, employee_name)
        bot.send_photo(message.chat.id, photo=open(filename_jpeg, 'rb'), caption=employee_name)
        #bot.send_document(message.chat.id, document =  open(filename, 'rb'))
        os.remove(filename_jpeg)
        return timesheet_buttons(message, employee_name)
    elif text == list_buttons[2]:
        msg = bot.send_message(message.chat.id, 'Введи имя')
        bot.register_next_step_handler(msg, edit_name)
    elif text == list_buttons[1] or back == "Назад":
        markup = types.ReplyKeyboardMarkup()
        button_back = types.KeyboardButton("Назад") 
        markup.add(button_back)    
        asddc = formatting.mbold("!!!Недопустимо использование пробелов в указании диапазона дат!!!")
        separator = f'Введи необходимые данные этого месяца, в формате:\n 1-5 ПГВР XXX.XX\nНезаполненные даты будут считаться работой в офисе 8 часов\n\n{formatting.escape_markdown(asddc)}\n\n\
Если объектов больше одного, перечисли через запятую\n\
Пример:"2-5 ПГВР 429.00, 6-13 ПГВР 123.70 Торгили, 14-20 отпуск, 21-31 больничный"'
        msg = bot.send_message(message.chat.id, formatting.escape_markdown(separator), reply_markup = markup, parse_mode='MarkdownV2')
        return bot.register_next_step_handler(msg, split_message, employee_name, list_buttons)
    elif text == "Выгрузить табель":
        locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
        now = datetime.datetime.now()
        year = now.year
        month = now.month
        filename_xlsx = f'Табель {calendar.month_name[month]  } {year}.xlsx'
        bot.send_document(message.chat.id, document =  open(filename_xlsx, 'rb'))
        return timesheet_buttons(message, employee_name)
    else:
        bot.send_message(message.chat.id, f'Нажми на кнопку {employee_name}')
        return timesheet_buttons(message, employee_name)

def split_message(message, employee_name, buttons):
    try:
        if message.text == 'Назад':
            bot.send_message(message.chat.id, f'Нажми на кнопку {employee_name}')
            return timesheet_buttons(message, employee_name)
        now = datetime.datetime.now()
        text = message.text
        list_range_name = []
        list_days_descwork = text.split(',')
        for j, days_name in enumerate(list_days_descwork):
            list_name_days = days_name.strip().split(' ', 1)
            range_split = list_name_days[0].split('-')
            work_hour_split_and_descr = days_name.split(':', 2)
            if len(list_name_days) <= 1:
                list_name_days.append("")
            if ":" in list_name_days[1]:
                list_name_days[1] = list_name_days[1][:list_name_days[1].find(':')].strip()
            if len(range_split) <= 1:
                range_split.append(range_split[0])
            if len(work_hour_split_and_descr) == 2:
                work_hour_split_and_descr.append("")
            elif len(work_hour_split_and_descr) == 1:
                work_hour_split_and_descr.extend(["", ""])
            for i in range_split:
                if i.isdigit() and int(i) in range(1,calendar.mdays[now.month] + 1) and int(range_split[0]) <= int(range_split[1]):
                    continue
                else:
                    raise Exception() 
            list_range_name.extend([[int(range_split[0]), 
                                            int(range_split[1]), 
                                            list_name_days[1]]]) 
            if list_name_days[1].lower() in ['отпуск','больничный',"отгул",""]: 
                list_range_name[j].append("")  
                list_range_name[j].append("") 
            else: 
                if work_hour_split_and_descr[1].strip().isdigit():
                    list_range_name[j].append(int(work_hour_split_and_descr[1]))
                else:
                    work_hour_split_and_descr[2] = work_hour_split_and_descr[1]
                    list_range_name[j].append(8)       
            list_range_name[j].append(work_hour_split_and_descr[2]) 
        filename = default_timesheet (employee_name = employee_name)
        edit_timesheet(filename, list_range_name, employee_name)
        #thr1 = threading.Thread(target = export_excel_jpeg(filename, employee_name)).start()
        filename_jpeg = export_excel_jpeg(filename, employee_name)
        bot.send_photo(message.chat.id, photo=open(filename_jpeg, 'rb'), caption=employee_name)
        bot.send_message(377190896, employee_name + ' .Сделал отчет!')
        #bot.send_document(message.chat.id, document =  open(filename_jpeg, 'rb'))
        os.remove(filename_jpeg)
        return timesheet_buttons(message, employee_name)
    except Exception as e:
            bot.send_message(message.chat.id,"Неверный формат ввода!\n\n\
Если перечисляешь разные объекты обязательно ставь между ними \",\"\n\
Пример:\"2-5 ПГВР 429.00, 6-13 ПГВР 123.70 Торгили, 14-20 отпуск, 21-31 больничный\"")
            #print(f"\n\n\n\n\n\{e}\n\n\n\n\n\n\n\n\n\n")
            bot.send_message(377190896, e)
            return get_timesheet(message, buttons, employee_name, back = 'Назад' )

def edit_name(message):
    return timesheet_buttons(message, message.text)

def timesheet(descript_work, employee_name):
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    line_shift = 3
    now = datetime.datetime.now()
    file_name = f'Табель {now.year} {calendar.month_name[now.month]}.xlsx'
    green_color = '#92d050'
    red_color = '#ff0000'
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet('Табель')
    format_cell = workbook.add_format({'bold': True,
                                'fg_color' :f'{green_color}',
                                'border':   1},
                                ) 
    worksheet.write(f'A1', "ФИО:", format_cell)
    worksheet.write(f'A2', "Месяц:", format_cell)
    worksheet.write(f'A3', "Дата", format_cell)
    worksheet.write(f'B1', employee_name, format_cell)
    worksheet.write(f'B2', f"{calendar.month_name[now.month]} {str(now.year)}", format_cell)
    worksheet.write(f'B3', "Описание выполненных работ", format_cell)
    format_cell.set_align('center')
    worksheet.write(f'C1', "", format_cell)
    worksheet.write(f'C2', "", format_cell)
    worksheet.write(f'C3', "Время работы, ч", format_cell)
    worksheet.write(f'D1', "", format_cell)
    worksheet.write(f'D2', "", format_cell)
    worksheet.write(f'D3', "Комментарии", format_cell)
    month_calendar = calendar.monthcalendar(now.year,now.month)
    asd = []
    list_name_day = list(calendar.day_abbr) 
    for list_week in month_calendar:
        calendar_month = {}
        for i, day in enumerate(list_week):
            calendar_month[list_name_day[i]] = day
        asd.append(calendar_month)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 15)
    for dict_week in asd:
        for name_day, num_day in dict_week.items():
            if name_day in ['Сб','Вс'] and num_day != 0:
                work_and_travel = ""
                num_work_hour = ""
                format_cell_date = workbook.add_format({'bold': True,
                                'fg_color' :f'{red_color}',
                                'border':   1},
                                ) 
                format_cell_work = workbook.add_format({'bold': True,
                                'fg_color' :f'{red_color}',
                                'border':   1},
                                ) 
                
            elif num_day != 0:
                work_and_travel = descript_work
                num_work_hour = 8
                format_cell_date = workbook.add_format({'bold': True,
                                'fg_color' :f'{green_color}',
                                'border':   1},
                                ) 
                format_cell_work = workbook.add_format({'bold': True,
                                'fg_color' :'#ffffff',
                                'border':   1},
                                ) 
            else:
                continue
            worksheet.write(f'A{num_day + line_shift}', name_day + " " + str(num_day), format_cell_date) 
            worksheet.write(f'B{num_day + line_shift}', work_and_travel, format_cell_work)
            worksheet.write(f'C{num_day + line_shift}', num_work_hour, format_cell_work)  
    workbook.close()
    return file_name

if __name__ == '__main__':
    while True:
        try:#добавляем try для бесперебойной работы
            bot.infinity_polling()
        except:
            time.sleep(10)#в случае падения

