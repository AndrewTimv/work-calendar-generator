# main file

import calendar

import yaml

from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Alignment, NamedStyle, Border, Side, PatternFill

months = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь',
]

days_of_week = [
    'пн',
    'вт',
    'ср',
    'чт',
    'пт',
    'сб',
    'вс',
]

def main():
    # Собираем год и фамилии
    with open('config.yaml') as f:
        config = yaml.safe_load(f)

    year = int(config['year'])
    persons = config['persons']

    # Создаём xls
    wb = Workbook()
    ws = wb.active

    # Создаём стили
    month_font = NamedStyle(name='month_font')
    month_font.font = Font(bold=True)
    bd = Side(style='thin', color="000000")
    month_font.border = Border(left=bd, top=bd, right=bd, bottom=bd)
    month_font.alignment = Alignment(horizontal='center')

    cell_style = NamedStyle(name='cell')
    cell_style.border = Border(left=bd, top=bd, right=bd, bottom=bd)

    # Счетчик строк
    row = 2

    # Сжимаем колонки
    for col in ws.iter_cols(2, 33):
        ws.column_dimensions[col[0].column_letter].width = 3


    # Для каждого месяца
    for i, month in enumerate(months):
        month_num = i + 1

        # Новый месяц
        calendar_month = calendar.monthrange(year, month_num)
        start_day = int(calendar_month[0])
        days_in_month = int(calendar_month[1])

        # Заголовок месяца
        print(month)
        month_cell = f'B{row}'
        ws[month_cell] = month
        ws[month_cell].style = month_font
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=1+days_in_month)
        row += 1


        n = 1
        day = start_day - 1
        for col in ws.iter_cols(2, days_in_month+1):
            # Числа
            print(f'{col[0].column_letter}{row}')
            ws[f'{col[0].column_letter}{row}'] = n
            ws[f'{col[0].column_letter}{row}'].style = cell_style
            n += 1
            # Дни недели
            ws[f'{col[0].column_letter}{row+1}'] = days_of_week[day % 7]
            ws[f'{col[0].column_letter}{row + 1}'].style = cell_style
            # Раскрашиваем выходные
            if day % 7 in [5,6]:
                ws[f'{col[0].column_letter}{row}'].fill = PatternFill('solid', '00FF9900')
                ws[f'{col[0].column_letter}{row+1}'].fill = PatternFill('solid', '00FF9900')
                for m in range(len(persons)):
                    ws[f'{col[0].column_letter}{row + 2 + m}'].fill = PatternFill('solid', '00FF9900')
            day += 1
        row += 2

        # Имена
        for person in persons:
            print(person)
            ws[f'A{row}'] = person
            ws[f'A{row}'].style = cell_style
            for col in ws.iter_cols(2, days_in_month + 1):
                ws[f'{col[0].column_letter}{row}'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
            row += 1



    wb.save('calendar.xls')



if __name__ == '__main__':
    main()
