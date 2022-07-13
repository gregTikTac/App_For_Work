from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe2.docx')

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
# лист exel страницы(активный)
sheet = file_for_work.active

list_SOG_and_OG = []
list_of_og = []
list_of_sog = []


def converting_exel_files_to_list_for_sog():
    for row in sheet.rows:
        list_of_sog.append(str(row[0].value))


def converting_exel_files_to_list_for_og():
    """Перевод значений эксель в список (второй столбец столбец)"""
    for row in sheet.rows:
        list_of_og.append(str(row[1].value))


def print_numb_and_values_list_SOG():
    for numb, values in enumerate(list_of_sog):
        print(f'{numb} {values}')


def print_numb_and_values_list_OG():
    for numb, values in enumerate(list_of_og):
        print(f'{numb} {values}')



def list_SOG_for_template():
    for num, values in enumerate(list_of_sog):
        if sog != num:
            continue
        else:
            list_SOG_and_OG.append(values)


def list_OG_for_template():
    for num, values in enumerate(list_of_og):
        if og != num:
            continue
        else:
            list_SOG_and_OG.append(values)




converting_exel_files_to_list_for_sog()
del list_of_sog[0]
del list_of_sog[-12:]
print_numb_and_values_list_SOG()
print('+' + '---------------' * 10 + '+')
sog = int(input('Введите номер СОГ: '))

list_SOG_for_template()
print('Запись добавлена')
print(list_SOG_and_OG)




converting_exel_files_to_list_for_og()
del list_of_og[0:]
print_numb_and_values_list_OG()
print('+' + '---------------' * 10 + '+')
og = int(input('Введите номер ОГ: '))
list_OG_for_template()
print('Запись добавлена')
print(list_SOG_and_OG)