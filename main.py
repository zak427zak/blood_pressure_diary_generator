from tkinter.tix import *
import datetime
from datetime import timedelta
import xlwt
import random
import math
import xlrd

# Инициализируем книгу
book = xlwt.Workbook(encoding="utf-8")

name_for_save = 'default.xls'

# Создаем лист в книге
sheet1 = book.add_sheet("Дневник АД", cell_overwrite_ok=True)

# Пишем заголовки в документ
sheet1.write(0,0, "Дата", xlwt.easyxf('font: bold on;align: wrap on,vert centre, horiz center;border: left thin,right thin,top thin,bottom thin;'))
sheet1.write(0,1, "Утро", xlwt.easyxf('font: bold on;align: wrap on,vert centre, horiz center;border: left thin,right thin,top thin,bottom thin;'))
sheet1.write(0,2, "Вечер", xlwt.easyxf('font: bold on;align: wrap on,vert centre, horiz center;border: left thin,right thin,top thin,bottom thin;'))

# Если нужно, чтобы вечерние были ниже, чем утренние
def lower_in_evening(total_days):
    rb = xlrd.open_workbook(name_for_save)
    sheet = rb.sheet_by_index(0)

    for x in range(total_days):
        if (sheet.cell(1+x,2).value != ''):
            temp_left = sheet.cell(1+x,1).value.split('/')
            temp_right = sheet.cell(1+x,2).value.split('/')
            if int(temp_left[0]) > int(temp_right[0]):
                dope = temp_right[0]
                temp_right[0] = temp_left[0]
                temp_left[0] = dope
                sheet1.write(1 + x, 1, (temp_left[0] + '/' + temp_left[1]), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))
                sheet1.write(1 + x, 2, (temp_right[0] + '/' + temp_right[1]), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

# Если нужно, чтобы утренние были ниже, чем вечерние
def lower_in_morning(total_days):
    rb = xlrd.open_workbook(name_for_save)
    sheet = rb.sheet_by_index(0)

    for x in range(total_days):
        if (sheet.cell(1+x,2).value != ''):
            temp_left = sheet.cell(1+x,1).value.split('/')
            temp_right = sheet.cell(1+x,2).value.split('/')
            if int(temp_left[0]) < int(temp_right[0]):
                dope = temp_right[0]
                temp_right[0] = temp_left[0]
                temp_left[0] = dope
                sheet1.write(1 + x, 1, (temp_left[0] + '/' + temp_left[1]), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))
                sheet1.write(1 + x, 2, (temp_right[0] + '/' + temp_right[1]), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

# Рандомим элементы списка со значениями
def randomOrder_key(element):
    return random()

def run_calc():

    start_date_f = start_date_value.get()
    start_date = start_date_f.split('.')
    start_date_true = datetime.date(int(start_date[2]),int(start_date[1]),int(start_date[0]))

    end_date_f = end_date_value.get()
    end_date = end_date_f.split('.')
    end_date_true = datetime.date(int(end_date[2]), int(end_date[1]), int(end_date[0]))

    total_days = int((str(end_date_true-start_date_true).split()[0])) + 1

    data_list = []

    first_col = sheet1.col(0)
    first_col.width = 3500
    second_col = sheet1.col(1)
    second_col.width = 3500
    third_col = sheet1.col(2)
    third_col.width = 3500

    for x in range(total_days):
        up_date = timedelta(x)
        sheet1.write(x+1 ,0,(start_date_true + up_date), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin\
                                     ', num_format_str='DD.MM.YYYY'))
    need_results = total_days * 2
    need_upper = int(need_results) * ((int(soot_value.get())) / 100)

    for x in range(math.ceil(need_upper)):
        upper_1 = random.randint(int(high_ad_upper_from_value.get()),int(high_ad_upper_to_value.get()))
        upper_2 = random.randint(int(high_ad_lower_from_value.get()),int(high_ad_lower_to_value.get()))
        data_list.append(str(upper_1) + ' / ' + str(upper_2))

    for x in range(need_results-(math.ceil(need_upper))):
        lower_1 = random.randint(int(low_ad_upper_from_value.get()), int(low_ad_upper_to_value.get()))
        lower_2 = random.randint(int(low_ad_lower_from_value.get()), int(low_ad_lower_to_value.get()))
        data_list.append(str(lower_1) + ' / ' + str(lower_2))

    random.shuffle(data_list)

    for x in range(total_days):
        sheet1.write(x+1,1, data_list.pop(), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

    for x in range(total_days):
        sheet1.write(x+1,2, data_list.pop(), xlwt.easyxf('align: wrap yes, vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'))

    global name_for_save
    name_for_save = your_name_value.get() + ' c ' + str(start_date_f) + ' по ' + str(end_date_f) + '.xls'
    book.save(name_for_save)
    if lang.get() == 2:
        lower_in_evening(total_days)
        book.save(name_for_save)
    elif lang.get() == 3:
        lower_in_morning(total_days)
        book.save(name_for_save)
    else:
        pass


def run_generator():
    pass

def collect_inserted_data(**kwargs):
    pass


def add_value_to_dict(label,value,values_dict):
    values_dict[label.cget('text')] = value.get()

def create_and_configure_window(normal_font, normal_width, small_width, values_dict):
    start_window = Tk()
    start_window.title("Генератор АД, v1.1")
    start_window.geometry('550x620')

    lbl_main = Label(start_window, text="Заполните предложенные поля, затем нажмите кнопку внизу", font=normal_font)
    lbl_main.place(x=20, y=40)


    your_name_label = Label(start_window, text='Введите Ваше имя:', width=normal_width)
    your_name_label.place(x=40, y=70)
    your_name_value = Entry(start_window, width=15)
    your_name_value.place(x=200, y=70)
    your_name_value.insert(0, 'Василий Иванов')

    add_value_to_dict(your_name_label, your_name_value, values_dict)


    lbl_periods = Label(start_window, text="1. Задайте периоды:", font=normal_font)
    lbl_periods.place(x=40, y=70)

    start_date_label = Label(start_window, text='Дата начала измерения:', width=normal_width)
    start_date_label.place(x=40, y=100)
    start_date_value = Entry(start_window, width=15)
    start_date_value.place(x=200, y=100)
    start_date_value.insert(0, '01.05.2019')
    add_value_to_dict(start_date_label, start_date_value, values_dict)

    end_date_label = Label(start_window, text='Дата конца измерения:', width=normal_width)
    end_date_label.place(x=40, y=130)
    end_date_value = Entry(start_window, width=15)
    end_date_value.place(x=200, y=130)
    end_date_value.insert(0, '01.08.2019')
    add_value_to_dict(end_date_label, end_date_value, values_dict)

    lbl_periods = Label(start_window, text="2. Укажите границы АД:", font=normal_font)
    lbl_periods.place(x=40, y=170)

    lbl_high_ad = Label(start_window, text="Для повышенного давления:", font=normal_font)
    lbl_high_ad.place(x=40, y=200)

    high_ad_upper_from_label = Label(start_window, text='Верхнее (первая цифра), от:', width=normal_width)
    high_ad_upper_from_label.place(x=40, y=230)
    high_ad_upper_from_value = Entry(start_window, width=small_width)
    high_ad_upper_from_value.place(x=220, y=230)
    high_ad_upper_from_value.insert(0, '138')
    add_value_to_dict(high_ad_upper_from_label, high_ad_upper_from_value, values_dict)
    
    high_ad_upper_to_label = Label(start_window, text=' до: ', width=small_width)
    high_ad_upper_to_label.place(x=260, y=230)
    high_ad_upper_to_value = Entry(start_window, width=small_width)
    high_ad_upper_to_value.place(x=310, y=230)
    high_ad_upper_to_value.insert(0, '157')
    add_value_to_dict(high_ad_upper_from_label, high_ad_upper_from_value, values_dict)

    high_ad_lower_from_label = Label(start_window, text='Нижнее (вторая цифра), от:', width=normal_width)
    high_ad_lower_from_label.place(x=40, y=260)
    high_ad_lower_from_value = Entry(start_window, width=small_width)
    high_ad_lower_from_value.place(x=220, y=260)
    high_ad_lower_from_value.insert(0, '85')

    high_ad_lower_to_label = Label(start_window, text=' до: ', width=small_width)
    high_ad_lower_to_label.place(x=260, y=260)
    high_ad_lower_to_value = Entry(start_window, width=small_width)
    high_ad_lower_to_value.place(x=310, y=260)
    high_ad_lower_to_value.insert(0, '99')

    lbl_low_ad = Label(start_window, text="Для нормального давления:", font=normal_font)
    lbl_low_ad.place(x=40, y=300)

    low_ad_upper_from_label = Label(start_window, text='Верхнее (первая цифра), от:', width=normal_width)
    low_ad_upper_from_label.place(x=40, y=330)
    low_ad_upper_from_value = Entry(start_window, width=small_width)
    low_ad_upper_from_value.place(x=220, y=330)
    low_ad_upper_from_value.insert(0, '125')

    low_ad_upper_to_label = Label(start_window, text=' до: ', width=small_width)
    low_ad_upper_to_label.place(x=260, y=330)
    low_ad_upper_to_value = Entry(start_window, width=small_width)
    low_ad_upper_to_value.place(x=310, y=330)
    low_ad_upper_to_value.insert(0, '132')

    low_ad_lower_from_label = Label(start_window, text='Нижнее (вторая цифра), от:', width=normal_width)
    low_ad_lower_from_label.place(x=40, y=360)
    low_ad_lower_from_value = Entry(start_window, width=small_width)
    low_ad_lower_from_value.place(x=220, y=360)
    low_ad_lower_from_value.insert(0, '80')

    low_ad_lower_to_label = Label(start_window, text=' до: ', width=small_width)
    low_ad_lower_to_label.place(x=260, y=360)
    low_ad_lower_to_value = Entry(start_window, width=small_width)
    low_ad_lower_to_value.place(x=310, y=360)
    low_ad_lower_to_value.insert(0, '85')

    lbl_add = Label(start_window, text="3. Прочие параметры:", font=normal_font)
    lbl_add.place(x=40, y=400)

    soot_label = Label(start_window, text='Повышенного давления, в %', width=normal_width)
    soot_label.place(x=40, y=430)
    soot_value = Entry(start_window, width=15)
    soot_value.place(x=220, y=430)
    soot_value.insert(0, '70')
    soot_label_after = Label(start_window, text='(остальное будет нормальным)', width=normal_width)
    soot_label_after.place(x=260, y=430)

    lang = IntVar()

    if_standart_checkbutton = Radiobutton(text="Простое заполнение (значения будут разбросаны в случайном порядке)",
                                          variable=lang, value=1)
    if_standart_checkbutton.place(x=40, y=460)

    if_need_night_lower_checkbutton = Radiobutton(
        text="Вечером должно быть ниже (вечерняя пара АД будет всегда ниже, чем утренняя)", value=2, variable=lang)
    if_need_night_lower_checkbutton.place(x=40, y=490)
    if_need_night_lower_checkbutton.select()

    if_need_morning_lower_checkbutton = Radiobutton(
        text="Утром должно быть ниже (утренняя пара АД будет всегда ниже, чем вечерняя)", value=3, variable=lang)
    if_need_morning_lower_checkbutton.place(x=40, y=520)

    to_generate_btn = Button(start_window, text="Сгенерировать АД", command=run_generator)
    to_generate_btn.place(x=40, y=560)

    start_window.mainloop()


def main():
    normal_font = ("Arial Bold", 11)
    
    normal_width = 25
    small_width = 5

    values_dict = {}

    create_and_configure_window(normal_font, normal_width, small_width, values_dict)


if __name__ == '__main__':
    main()