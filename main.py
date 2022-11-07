import datetime
import math
import random
from datetime import datetime
from datetime import timedelta
from functools import partial
from tkinter import Entry, Label, Button, Radiobutton, Tk, IntVar

import xlrd
import xlwt


def lower_in_evening(total_days, current_list, name_for_save):
    rb = xlrd.open_workbook(name_for_save)
    sheet = rb.sheet_by_index(0)

    for x in range(total_days):
        if sheet.cell(1 + x, 2).value != '':
            temp_left = sheet.cell(1 + x, 1).value.split('/')
            temp_right = sheet.cell(1 + x, 2).value.split('/')
            if int(temp_left[0]) > int(temp_right[0]):
                dope = temp_right[0]
                temp_right[0] = temp_left[0]
                temp_left[0] = dope
                current_list.write(1 + x, 1, (temp_left[0] + '/' + temp_left[1]), xlwt.easyxf(
                    'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))
                current_list.write(1 + x, 2, (temp_right[0] + '/' + temp_right[1]), xlwt.easyxf(
                    'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))


def lower_in_morning(total_days, current_list, name_for_save):
    rb = xlrd.open_workbook(name_for_save)
    sheet = rb.sheet_by_index(0)

    for x in range(total_days):
        if sheet.cell(1 + x, 2).value != '':
            temp_left = sheet.cell(1 + x, 1).value.split('/')
            temp_right = sheet.cell(1 + x, 2).value.split('/')
            if int(temp_left[0]) < int(temp_right[0]):
                dope = temp_right[0]
                temp_right[0] = temp_left[0]
                temp_left[0] = dope
                current_list.write(1 + x, 1, (temp_left[0] + '/' + temp_left[1]), xlwt.easyxf(
                    'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))
                current_list.write(1 + x, 2, (temp_right[0] + '/' + temp_right[1]), xlwt.easyxf(
                    'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))


def run_generator(values_dict):
    current_book, current_list = prepare_final_list()

    start_date_true = datetime.strptime(values_dict['Дата начала измерения:'], '%d.%m.%Y')
    end_date_true = datetime.strptime(values_dict['Дата конца измерения:'], '%d.%m.%Y')
    total_days = int((str(end_date_true - start_date_true).split()[0])) + 1

    data_list = []

    for x in range(total_days):
        up_date = timedelta(x)
        current_list.write(x + 1, 0, (start_date_true + up_date), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin\
                                     ', num_format_str='DD.MM.YYYY'))
    need_results = total_days * 2
    need_upper = int(need_results) * ((int(values_dict['Повышенного давления, в %'])) / 100)

    for x in range(math.ceil(need_upper)):
        upper_1 = random.randint(int(values_dict['Для повышенного давления: Верхнее (первая цифра), от:']),
                                 int(values_dict['Для повышенного давления:  Верхнее, до: ']))
        upper_2 = random.randint(int(values_dict['Для повышенного давления: Нижнее (вторая цифра), от:']),
                                 int(values_dict['Для повышенного давления:  Нижнее, до: ']))
        data_list.append(str(upper_1) + ' / ' + str(upper_2))

    for x in range(need_results - (math.ceil(need_upper))):
        lower_1 = random.randint(int(values_dict['Для нормального давления: Верхнее (первая цифра), от:']),
                                 int(values_dict['Для нормального давления:  Верхнее, до: ']))
        lower_2 = random.randint(int(values_dict['Для нормального давления: Нижнее (вторая цифра), от:']),
                                 int(values_dict['Для нормального давления:  Нижнее, до: ']))
        data_list.append(str(lower_1) + ' / ' + str(lower_2))

    random.shuffle(data_list)

    for x in range(total_days):
        current_list.write(x + 1, 1, data_list.pop(), xlwt.easyxf(
            'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

    for x in range(total_days):
        current_list.write(x + 1, 2, data_list.pop(), xlwt.easyxf(
            'align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

    name_for_save = values_dict['Введите Ваше имя:'] + ' c ' + str(start_date_true.date()) + ' по ' + str(
        end_date_true.date()) + '.xls'
    current_book.save(name_for_save)
    if values_dict['Условия заполнения'] == 2:
        lower_in_evening(total_days, current_list, name_for_save)
        current_book.save(name_for_save)
    elif values_dict['Условия заполнения'] == 3:
        lower_in_morning(total_days, current_list, name_for_save)
        current_book.save(name_for_save)
    print('Готово!')


def get_data_to_generate(value, values_dict, *args):
    label = ''
    for a in args:
        label += a.cget('text')
    values_dict[label] = value.get()


def prepare_final_list():
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Дневник АД", cell_overwrite_ok=True)
    normal_style = xlwt.easyxf(
        'font: bold on;align: wrap on,vert centre, horiz center;border: left thin,right thin,top thin,bottom thin;')

    # Пишем заголовки в документ
    sheet1.write(0, 0, "Дата", normal_style)
    sheet1.write(0, 1, "Утро", normal_style)
    sheet1.write(0, 2, "Вечер", normal_style)

    first_col = sheet1.col(0)
    first_col.width = 3500
    second_col = sheet1.col(1)
    second_col.width = 3500
    third_col = sheet1.col(2)
    third_col.width = 3500

    return book, sheet1


def create_and_configure_window(normal_font, normal_width, small_width, values_dict):
    start_window = Tk()
    start_window.title("Генератор АД, v1.1 от 07.11.2019")
    start_window.geometry('550x700')

    start_y = 40
    start_x = 20

    # Title block
    lbl_main = Label(start_window, text="Заполните предложенные поля, затем нажмите кнопку внизу", font=normal_font)
    lbl_main.place(x=start_x, y=start_y)

    height = 30
    your_name_label = Label(start_window, text='Введите Ваше имя:', width=normal_width)
    your_name_label.place(x=25, y=start_y + height)
    your_name_value = Entry(start_window)
    your_name_value.place(x=200, y=70)
    your_name_value.insert(0, 'Василий Иванов')

    get_data_to_generate(your_name_value, values_dict, your_name_label)

    # lbl_periods = Label(start_window, text="1. Задайте периоды:", font=normal_font)
    # lbl_periods.place(x=40, y=70)

    start_date_label = Label(start_window, text='Дата начала измерения:', width=normal_width)
    start_date_label.place(x=40, y=100)
    start_date_value = Entry(start_window, width=15)
    start_date_value.place(x=200, y=100)
    start_date_value.insert(0, datetime.strftime(datetime.today(), '%d.%m.%Y'))
    get_data_to_generate(start_date_value, values_dict, start_date_label)

    end_date_label = Label(start_window, text='Дата конца измерения:', width=normal_width)
    end_date_label.place(x=36, y=130)
    end_date_value = Entry(start_window, width=15)
    end_date_value.place(x=200, y=130)
    end_date_value.insert(0, datetime.strftime(datetime.today() + timedelta(days=30), '%d.%m.%Y'))
    get_data_to_generate(end_date_value, values_dict, end_date_label)

    lbl_periods = Label(start_window, text="1. Укажите границы АД:", font=normal_font)
    lbl_periods.place(x=40, y=170)

    lbl_high_ad = Label(start_window, text="Для повышенного давления: ", font=normal_font)
    lbl_high_ad.place(x=40, y=200)

    high_ad_upper_from_label = Label(start_window, text='Верхнее (первая цифра), от:', width=normal_width)
    high_ad_upper_from_label.place(x=40, y=230)
    high_ad_upper_from_value = Entry(start_window, width=small_width)
    high_ad_upper_from_value.place(x=220, y=230)
    high_ad_upper_from_value.insert(0, '138')
    get_data_to_generate(high_ad_upper_from_value, values_dict, lbl_high_ad, high_ad_upper_from_label)

    high_ad_upper_to_label = Label(start_window, text=' Верхнее, до: ', width=normal_width, anchor="w")
    high_ad_upper_to_label.place(x=260, y=230)
    high_ad_upper_to_value = Entry(start_window, width=small_width)
    high_ad_upper_to_value.place(x=350, y=230)
    high_ad_upper_to_value.insert(0, '157')
    get_data_to_generate(high_ad_upper_to_value, values_dict, lbl_high_ad, high_ad_upper_to_label)

    high_ad_lower_from_label = Label(start_window, text='Нижнее (вторая цифра), от:', width=normal_width)
    high_ad_lower_from_label.place(x=40, y=260)
    high_ad_lower_from_value = Entry(start_window, width=small_width)
    high_ad_lower_from_value.place(x=220, y=260)
    high_ad_lower_from_value.insert(0, '85')
    get_data_to_generate(high_ad_lower_from_value, values_dict, lbl_high_ad, high_ad_lower_from_label)

    high_ad_lower_to_label = Label(start_window, text=' Нижнее, до: ', width=normal_width, anchor="w")
    high_ad_lower_to_label.place(x=260, y=260)
    high_ad_lower_to_value = Entry(start_window, width=small_width)
    high_ad_lower_to_value.place(x=350, y=260)
    high_ad_lower_to_value.insert(0, '99')
    get_data_to_generate(high_ad_lower_to_value, values_dict, lbl_high_ad, high_ad_lower_to_label)

    lbl_low_ad = Label(start_window, text="Для нормального давления: ", font=normal_font)
    lbl_low_ad.place(x=40, y=300)

    low_ad_upper_from_label = Label(start_window, text='Верхнее (первая цифра), от:', width=normal_width)
    low_ad_upper_from_label.place(x=40, y=330)
    low_ad_upper_from_value = Entry(start_window, width=small_width)
    low_ad_upper_from_value.place(x=220, y=330)
    low_ad_upper_from_value.insert(0, '125')
    get_data_to_generate(low_ad_upper_from_value, values_dict, lbl_low_ad, low_ad_upper_from_label)

    low_ad_upper_to_label = Label(start_window, text=' Верхнее, до: ', width=normal_width, anchor="w")
    low_ad_upper_to_label.place(x=260, y=330)
    low_ad_upper_to_value = Entry(start_window, width=small_width)
    low_ad_upper_to_value.place(x=350, y=330)
    low_ad_upper_to_value.insert(0, '132')
    get_data_to_generate(low_ad_upper_to_value, values_dict, lbl_low_ad, low_ad_upper_to_label)

    low_ad_lower_from_label = Label(start_window, text='Нижнее (вторая цифра), от:', width=normal_width)
    low_ad_lower_from_label.place(x=40, y=360)
    low_ad_lower_from_value = Entry(start_window, width=small_width)
    low_ad_lower_from_value.place(x=220, y=360)
    low_ad_lower_from_value.insert(0, '80')
    get_data_to_generate(low_ad_lower_from_value, values_dict, lbl_low_ad, low_ad_lower_from_label)

    low_ad_lower_to_label = Label(start_window, text=' Нижнее, до: ', width=normal_width, anchor="w")
    low_ad_lower_to_label.place(x=260, y=360)
    low_ad_lower_to_value = Entry(start_window, width=small_width)
    low_ad_lower_to_value.place(x=350, y=360)
    low_ad_lower_to_value.insert(0, '85')
    get_data_to_generate(low_ad_lower_to_value, values_dict, lbl_low_ad, low_ad_lower_to_label)

    lbl_add = Label(start_window, text="2. Прочие параметры:", font=normal_font)
    lbl_add.place(x=40, y=400)

    soot_label = Label(start_window, text='Повышенного давления, в %', width=normal_width)
    soot_label.place(x=40, y=430)
    soot_value = Entry(start_window, width=15)
    soot_value.place(x=220, y=430)
    soot_value.insert(0, '70')
    get_data_to_generate(soot_value, values_dict, soot_label)
    soot_label_after = Label(start_window, text='(остальное будет нормальным)', width=normal_width)
    soot_label_after.place(x=260, y=430)

    lbl_add_condit = Label(start_window, text="Условия заполнения", font=normal_font)
    lbl_add_condit.place(x=40, y=490)

    lang = IntVar()

    if_standart_checkbutton = Radiobutton(text="Простое заполнение (значения будут разбросаны в случайном порядке)",
                                          variable=lang, value=1)
    if_standart_checkbutton.place(x=40, y=530)

    if_need_night_lower_checkbutton = Radiobutton(
        text="Вечером должно быть ниже (вечерняя пара АД будет всегда ниже, чем утренняя)", value=2, variable=lang)
    if_need_night_lower_checkbutton.place(x=40, y=570)
    if_need_night_lower_checkbutton.select()

    if_need_morning_lower_checkbutton = Radiobutton(
        text="Утром должно быть ниже (утренняя пара АД будет всегда ниже, чем вечерняя)", value=3, variable=lang)
    if_need_morning_lower_checkbutton.place(x=40, y=610)

    get_data_to_generate(lang, values_dict, lbl_add_condit)

    to_generate_btn = Button(start_window, text="Сгенерировать АД", command=partial(run_generator, values_dict))
    to_generate_btn.place(x=40, y=650)

    start_window.mainloop()


def main():
    normal_font = ("Arial Bold", 11)
    normal_width = 25
    small_width = 5
    values_dict = {}
    create_and_configure_window(normal_font, normal_width, small_width, values_dict)


if __name__ == '__main__':
    main()
