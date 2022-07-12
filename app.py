from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe2.docx')

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
# лист exel страницы(активный)
sheet = file_for_work.active


def converting_exel_files_to_list():
    """Перевод значений эксель в список"""
    for row in sheet.rows:
        lst.append(str(row[0].value))


def converting_data_lst_to_dict_data():
    """преобразует список в словрь"""
    # dict = {index: value for index, value in enumerate(lst} ---> то же самое
    for index, value in enumerate(lst):
        dict_of_lst_og[index] = value


def pretty_print_to_dict():
    """Выводит словарь ключ-значение-построчно"""
    for key, value in dict_of_lst_og.items():
        print(f'{key}: {value}')


def change_sog_of_record():
    """Позволяет менять л/с в записи"""
    try:
        counter = 0
        while True:
            if sog == counter:
                context = {
                    'sog': dict_of_lst_og[counter]
                }
                doc.render(context)
                doc.save('probe2.docx')
                break
            else:
                counter += 1
    except KeyError:
        print("Введите значение в указанном диапазоне")


lst = []
converting_exel_files_to_list()
del lst[0]
del lst[-12:]

dict_of_lst_og = {}
converting_data_lst_to_dict_data()

pretty_print_to_dict()

print('+' + '-------' * 10 + '+')
sog = int(input("Выберите Старшего(нажмите соотвествующую цифру): "))
change_sog_of_record()