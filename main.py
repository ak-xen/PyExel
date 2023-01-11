import openpyxl
import pandas as pd

cols = [10, 11, 13, 15, 19, 20, 24, 27]
wb = pd.read_excel('from.XLSX', engine='openpyxl', usecols=cols)
type_obj = ['Пусковая котельная',
            'Установка ХВО Пусковой котельной',
            'Котельная Расширение',
            'Обратноосмотическая установка',
            'Установка очистки парового конденсата']
list_type_obj = []
for i, row in enumerate(wb.values):
    if len(set(row)) == 1:
        continue

    obj = row[0]
    if obj in type_obj:
        list_type_obj.append(row)

print(f'Длина всего списка = {len(list_type_obj)}')
count_records = 0
pipelines = ['трубопровод', 'дренажи', 'паропровод', 'труб-д']
pumps = ['насос', 'электронасосный']
reservoirs = ['бак', 'рдвд', 'рднд', 'бак-мерник']
vessels = ['котел', 'фильтра', 'деаэратор', 'сепаратор', 'подогреватель', 'фильтр', 'теплообменный']

list_pipe = []
list_pump = []
list_res = []
list_ves = []
list_other = []


def val(string, value):
    for eq, l in (pipelines, list_pipe), (pumps, list_pump), (reservoirs, list_res), (vessels, list_ves):
        for st in eq:
            if st in string:
                l.append(list(value))
                return True
    return False


for value in list_type_obj:
    count_records += 1
    equip = value[2]
    equip = equip.strip().lower()
    objct, type_obj, eq, type_rep, start, end, data, other = value
    vals = [objct, eq, type_rep, type_obj, start, end, data, other]
    if val(equip, vals):
        continue
    else:
        list_other.append(value)


def cre(name, ls):
    wb = openpyxl.Workbook()

    for sheet in wb.worksheets:
        for data in ls:
            sheet.append(data)

    wb.save(f'{name}.xlsx')


for name, ls in ('pump', list_pump), ('pipe', list_pipe), ('res', list_res), ('ves', list_ves):
    cre(name, ls)
