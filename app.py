from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe2.docx')

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
# лист exel страницы(активный)
sheet = file_for_work.active


def converting_exel_files_to_list_for_sog():
    """Перевод значений эксель в список (первый столбец)"""
    for row in sheet.rows:
        list_of_sog.append(str(row[0].value))


def converting_exel_files_to_list_for_og():
    """Перевод значений эксель в список (второй столбец столбец)"""
    for row in sheet.rows:
        list_of_og.append(str(row[1].value))


def converting_sog_list_data_to_sog_dict_data():
    """преобразует список в словарь для старших"""
    # dict = {index: value for index, value in enumerate(lst} ---> то же самое
    for index, value in enumerate(list_of_sog):
        dict_of_sog[index] = value


def converting_og_list_data_to_og_dict_data():
    """преобразует список в словарь для состава управления"""
    # dict = {index: value for index, value in enumerate(lst} ---> то же самое
    for index, value in enumerate(list_of_og):
        dict_of_og[index] = value


def pretty_print_to_dict_sog():
    """Выводит словарь ключ-значение-построчно"""
    for key, value in dict_of_sog.items():
        print(f'{key}: {value}')


def pretty_print_to_dict_og():
    """Выводит словарь ключ-значение-построчно"""
    for key, value in dict_of_og.items():
        print(f'{key}: {value}')


def change_sog_of_record():
    """Позволяет менять л/с сог в записи"""
    try:
        counter = 0
        while True:
            if sog == counter:
                context = {
                    'sog': dict_of_sog[counter]
                }
                doc.render(context)
                doc.save('probe2.docx')
                break
            else:
                counter += 1
    except KeyError:
        print("Введите значение в указанном диапазоне")


def change_og_of_record():
    """Позволяет менять л/с ог в записи"""
    try:
        counter = 0
        while True:
            if og_upravlenie == counter:
                context = {
                    'sog': dict_of_og[counter]
                }
                doc.render(context)
                doc.save('probe2.docx')
                break
            else:
                counter += 1
    except KeyError:
        print("Введите значение в указанном диапазоне")


# # добавление старшего ог
list_of_sog = []
converting_exel_files_to_list_for_sog()
del list_of_sog[0]
del list_of_sog[-12:]

dict_of_sog = {}
converting_data_lst_to_dict_data()

pretty_print_to_dict()

print('+' + '-------' * 10 + '+')
sog = int(input("Выберите Старшего(нажмите соотвествующую цифру): "))
change_sog_of_record()


# добавление ог управления
list_of_og = []
converting_exel_files_to_list_for_og()

dict_of_og = {}
converting_og_list_data_to_og_dict_data()

pretty_print_to_dict_og()

print('+' + '-------' * 10 + '+')
og_upravlenie = int(input("Выберите off управления(нажмите соотвествующую цифру): "))