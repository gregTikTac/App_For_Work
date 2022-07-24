from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe3.docx')

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
# лист exel страницы(активный)
sheet = file_for_work.active
# л/с ОГ, ПАТ шаблон
list_SOG_and_OG_for_template = []
list_for_PAT_for_template = []
# ПМ, АК шаблон
list_of_PM_for_OG_for_template = []
list_of_AK_for_PAT_for_template = []
# л/с
list_of_og = []
list_of_sog = []
list_of_pat = []
list_PM = []
list_AK = []



def converting_exel_files_to_list_for_sog():
    """Конвертирует данные из 1 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_of_sog.append(str(row[0].value))


def converting_exel_files_to_list_for_og():
    """Перевод значений эксель в список (второй столбец столбец)"""
    for row in sheet.rows:
        list_of_og.append(str(row[1].value))


def converting_exel_files_to_list_for_PAT():
    """Конвертирует данные из 3 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_of_pat.append(str(row[2].value))


def converting_exel_files_to_list_PM_for_PAT():
    """Конвертирует данные из 4 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_PM.append(str(row[3].value))


def converting_exel_files_to_list_AK_for_PAT():
    """Конвертирует данные из 5 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_AK.append(str(row[4].value))


def print_numb_and_values_list_SOG():
    """Вывод списка [номер : значение]"""
    print("Список СОГ:")
    for numb, values in enumerate(list_of_sog):
        print(f'{numb} {values}')
    print('+' + '---------------' * 10 + '+')


def print_numb_and_values_list_OG():
    """Вывод списка [номер : значение]"""
    print("Список ОГ:")
    for numb, values in enumerate(list_of_og):
        print(f'{numb} {values}')
    print('+' + '---------------' * 10 + '+')


def print_numb_and_values_list_PAT():
    """Вывод списка [номер : значение]"""
    print("Список ПАТ:")
    for numb, values in enumerate(list_of_pat):
        print(f'{numb} {values}')
    print('+' + '---------------' * 10 + '+')


def print_numb_and_values_list_PM_PAT():
    """Вывод списка [номер : значение]"""
    print("Список ПМ ОГ:")
    for numb, values in enumerate(list_PM):
        print(f'{numb} {values}')
    print('+' + '---------------' * 10 + '+')


def print_numb_and_values_list_AK_PAT():
    """Вывод списка [номер : значение]"""
    print("Список АК ПАТ:")
    for numb, values in enumerate(list_AK):
        print(f'{numb} {values}')
    print('+' + '---------------' * 10 + '+')


def list_SOG_for_template(number_of_sog):
    """Добавление данных выбранных данных СОГ в список для рендеренга страницы"""
    counter = 0
    while True:
        try:
            number_of_sog = int(number_of_sog)
            if number_of_sog < 0:
                print("Вы ввели не правильное значение")
                break
            elif counter == number_of_sog:
                list_SOG_and_OG_for_template.append(list_of_sog[counter])
                print('Запись добавлена!')
                print('+' + '---------------' * 10 + '+')
                break
            else:
                counter += 1
        except (IndexError, ValueError):
            print("Вы ввели не правильное значение!")
            break


def list_OG_for_template(number_of_og):
    """Добавление данных выбранных данных ОГ в список для рендеренга страницы"""
    counter = 0
    while True:
        try:
            number_of_og = int(number_of_og)
            if number_of_og < 0:
                print("Вы ввели не правильное значение")
                break
            elif counter == number_of_og:
                list_SOG_and_OG_for_template.append(list_of_og[counter])
                print('Запись добавлена!')
                print('+' + '---------------' * 10 + '+')
                break
            else:
                counter += 1
        except (IndexError, ValueError):
            print("Вы ввели не правильное значение!")
            break


def list_PAT_for_template(number_of_pat):
    """Добавление данных выбранных данных ОГ в список для рендеренга страницы"""
    counter = 0
    while True:
        try:
            number_of_pat = int(number_of_pat)
            if number_of_pat < 0:
                print("Вы ввели не правильное значение")
                break
            elif counter == number_of_pat:
                list_for_PAT_for_template.append(list_of_pat[counter])
                print('Запись добавлена!')
                print('+' + '---------------' * 10 + '+')
                break
            else:
                counter += 1
        except (IndexError, ValueError):
            print("Вы ввели не правильное значение!")
            break

def list_PM_for_template(number_of_PM):
    """Добавление данных выбранных данных ОГ в список для рендеренга страницы"""
    counter = 0
    while True:
        try:
            number_of_PM = int(number_of_PM)
            if number_of_PM < 0:
                print("Вы ввели не правильное значение")
                break
            elif counter == number_of_PM:
                list_of_PM_for_OG_for_template.append(list_PM[counter])
                print('Запись добавлена!')
                print('+' + '---------------' * 10 + '+')
                break
            else:
                counter += 1
        except (IndexError, ValueError):
            print("Вы ввели не правильное значение!")
            break

def list_AK_for_template(number_of_AK):
    """Добавление данных выбранных данных ОГ в список для рендеренга страницы"""
    counter = 0
    while True:
        try:
            number_of_AK = int(number_of_AK)
            if number_of_AK < 0:
                print("Вы ввели не правильное значение")
                break
            elif counter == number_of_AK:
                list_of_AK_for_PAT_for_template.append(list_AK[counter])
                print('Запись добавлена!')
                print('+' + '---------------' * 10 + '+')
                break
            else:
                counter += 1
        except (IndexError, ValueError):
            print("Вы ввели не правильное значение!")
            break



def choose_person_from_SOG():
    list_SOG_for_template(input('Введите номер СОГ: '))
    if len(list_SOG_and_OG_for_template) > 0:
        print(list_SOG_and_OG_for_template)
        print('+' + '---------------' * 10 + '+')

    else:
        while len(list_SOG_and_OG_for_template) < 1:
            list_SOG_for_template(input("Введите корректное значение: "))
            if len(list_SOG_and_OG_for_template) >= 1:
                print(list_SOG_and_OG_for_template)
                print('+' + '---------------' * 10 + '+')


def choose_person_from_OG():
    list_OG_for_template(int(input('Введите номер ОГ: ')))
    if len(list_SOG_and_OG_for_template) > 0:
        print(list_SOG_and_OG_for_template)
        print('+' + '---------------' * 10 + '+')
    else:
        while len(list_SOG_and_OG_for_template) < 1:
            list_OG_for_template(input("Введите корректное значение: "))
            if len(list_SOG_and_OG_for_template) >= 1:
                print(list_SOG_and_OG_for_template)
                print('+' + '---------------' * 10 + '+')


def choose_person_from_PAT():
    list_PAT_for_template(int(input('Введите номер в/с для ПАТ: ')))
    if len(list_for_PAT_for_template) > 0:
        print(list_for_PAT_for_template)
        print('+' + '---------------' * 10 + '+')
    else:
        while len(list_for_PAT_for_template) < 1:
            list_PAT_for_template(input("Введите корректное значение: "))
            if len(list_for_PAT_for_template) >= 1:
                print(list_for_PAT_for_template)
                print('+' + '---------------' * 10 + '+')


def choose_PM_from_OG():
    list_PM_for_template(int(input('Введите номер ПМ для ОГ: ')))
    if len(list_of_PM_for_OG_for_template) > 0:
        print(list_of_PM_for_OG_for_template)
        print('+' + '---------------' * 10 + '+')
    else:
        while len(list_of_PM_for_OG_for_template) < 1:
            list_PAT_for_template(input("Введите корректное значение: "))
            if len(list_of_PM_for_OG_for_template) >= 1:
                print(list_of_PM_for_OG_for_template)
                print('+' + '---------------' * 10 + '+')

def choose_AK_from_PAT():
    list_AK_for_template(int(input('Введите номер AK для ОГ: ')))
    if len(list_of_AK_for_PAT_for_template) > 0:
        print(list_of_AK_for_PAT_for_template)
        print('+' + '---------------' * 10 + '+')
    else:
        while len(list_of_AK_for_PAT_for_template) < 1:
            list_PAT_for_template(input("Введите корректное значение: "))
            if len(list_of_AK_for_PAT_for_template) >= 1:
                print(list_of_AK_for_PAT_for_template)
                print('+' + '---------------' * 10 + '+')


def render_template():
# dict = {index: value for index, value in enumerate(lst}
    context = {
            'og0': list_SOG_and_OG_for_template[0],
            'og1': list_SOG_and_OG_for_template[1],
            'og2': list_SOG_and_OG_for_template[2],
            'og3': list_SOG_and_OG_for_template[3],
            'og4': list_SOG_and_OG_for_template[4],
            'og5': list_SOG_and_OG_for_template[5],
            'og6': list_SOG_and_OG_for_template[6],
            'og7': list_SOG_and_OG_for_template[7],
            'og8': list_SOG_and_OG_for_template[8],
        }

    doc.render(context)
    doc.save('probe3.docx')




    # content_PAT = {str(index): value for index, value in enumerate(list_for_PAT_for_template)}
    # content_PM = {str(index): value for index, value in enumerate(list_of_PM_for_OG_for_template)}
    # content_AK = {str(index): value for index, value in enumerate(list_of_AK_for_PAT_for_template)}

    # doc.render(content_PAT)
    # doc.render(content_PM)
    # doc.render(content_AK)




# ВЫБОР ОГ СОКРАЩЕННОГО СОСТАВА
converting_exel_files_to_list_for_sog()
del list_of_sog[0]
del list_of_sog[-32:]
print_numb_and_values_list_SOG()
choose_person_from_SOG()

converting_exel_files_to_list_for_og()
del list_of_og[0]
del list_of_og[-20:]
print_numb_and_values_list_OG()
choose_person_from_OG()

#
# # ВЫБОР ОГ ПОЛНОГО СОСТАВА
print('+' + '---------------' * 10 + '+')
print('Выберите ОГ полного состава')
print('+' + '---------------' * 10 + '+')
print()
choose_person_from_SOG()
for chose in range(0, 6):
    choose_person_from_OG()
print(list_SOG_and_OG_for_template)
#
# print()
# print('+' + '---------------' * 10 + '+')
# print("Выберете ПАТ")
# print('+' + '---------------' * 10 + '+')
# #
# # # ВЫБОР Л/С ДЛЯ ПАТ
# converting_exel_files_to_list_for_PAT()
# del list_of_pat[0]
# del list_of_pat[-15:]
# print_numb_and_values_list_PAT()
# choose_person_from_PAT()
# for choose in range(0, 24):
#     choose_person_from_PAT()
# #
# #
# print()
# print('+' + '---------------' * 10 + '+')
# print("Выберете ПМ для ОГ")
# print('+' + '---------------' * 10 + '+')
#
# #
# # # ВЫбор ПМ для ОГ
# converting_exel_files_to_list_PM_for_PAT()
# del list_PM[0]
# del list_PM[-2:]
# print_numb_and_values_list_PM_PAT()
# for choose in range(0, 8):
#     choose_PM_from_OG()
#
#
# print()
# print('+' + '---------------' * 10 + '+')
# print("Выберете AK для ПАТ")
# print('+' + '---------------' * 10 + '+')
#
# converting_exel_files_to_list_AK_for_PAT()
# del list_AK[-8:]
# print_numb_and_values_list_AK_PAT()
# for choose in range(0, 27):
#     choose_AK_from_PAT()


render_template()
