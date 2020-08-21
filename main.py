from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Color
from openpyxl.styles.borders import Border, Side
import numpy as np
import re
from data_preparation import run
from span import Span
from support import Support


def make_list(spans_list: list):
    k = 0
    supports_list = []
    end_index_list = []
    for x, y in spans_list:
        if k != 0 and k != x:
            end_index_list.append(int(k))
        if x != k:
            supports_list.append(int(x))
            end_index_list.append(int(x))
        supports_list.append(int(y))
        k = y
    end_index_list.append(int(k))
    return (supports_list, end_index_list)


def end_flag(support, arr):
    for item in arr:
        if item == support:
            return True
    return False


# открываем файл и берем текущий лист
filename = '1268 3'
full_filename = filename + '.xlsx'
workbook = load_workbook(full_filename)

# подготавливаем список из файла
run(full_filename, workbook)

# загружаем отформатированный список Excel
workbook = load_workbook(full_filename)
formatted_sheet = workbook['отформатированный']

# номер КТП (ST - substation transformer)
val_A2 = formatted_sheet['A2'].value
ST_number = int(re.findall(r'\d+', val_A2)[0])

# номер линии
val_B2 = formatted_sheet['B2'].value
line_number = int(re.findall(r'\d+', val_B2)[0])

# список пролётов
list_of_span_support_numbers = []
col_C = formatted_sheet['C']
for item in col_C:
    if item.value is None:
        break
    mFormat = re.findall(r'\d+', item.value)
    if mFormat:
        list_of_span_support_numbers.append(mFormat)

# список опор (вызываем функцию make_supports_list)
list_of_supports = make_list(list_of_span_support_numbers)[0]
end_supports_list = make_list(list_of_span_support_numbers)[1]

number_of_supports = len(list_of_supports)
number_of_spans = len(list_of_span_support_numbers)

list_of_marks = []
list_of_length = []
list_of_lamps = []
for i in range(0, number_of_spans):
    d = formatted_sheet.cell(row=i + 1, column=4)
    e = formatted_sheet.cell(row=i + 1, column=5)
    font_big = formatted_sheet.cell(row=i + 2, column=6)
    list_of_marks.append(d.value)
    list_of_length.append(round(e.value))
    list_of_lamps.append(font_big.value)

# создаем список с объектами пролетов
spans = []
for mark, span, length, lamp in zip(list_of_marks, list_of_span_support_numbers, list_of_length, list_of_lamps):
    flag = False
    if lamp == 'фонарка':
        flag = True
    span = Span(mark, span, length, flag)
    spans.append(span)

# создаём список с объектами опор
supports = []
temp_span = None
for cur_support in list_of_supports:
    for cur_span in spans:
        if end_flag(cur_support, end_supports_list):
            if int(cur_span.span_support1_number) == cur_support or int(cur_span.span_support2_number) == cur_support:
                support = Support(cur_support, cur_span)
                supports.append(support)
                break
        else:
            if int(cur_span.span_support2_number) == cur_support:
                temp_span = cur_span
                continue
            elif int(cur_span.span_support1_number) == cur_support:
                support = Support(cur_support, temp_span, cur_span)
                supports.append(support)
                break


# функции форматирования ячеек
def cell_border(row, col, side):
    my_ws.cell(row=row, column=col).border = my_ws.cell(row=row, column=col).border + side


def set_cell_height(row, height):
    rd = my_ws.row_dimensions[row]  # get dimension for row 3
    rd.height = height


def set_cell_width(col, width):
    rd = my_ws.column_dimensions[col]  # get dimension for row 3
    rd.width = width


def merge_cells_of_support(first_row, support):
    offset1 = 0
    offset2 = 0
    if support.span1.lamp is True:
        offset1 = 1
    nRow = 1 + offset1
    if support.span2:
        if support.span2.lamp is True:
            offset2 = 1
        nRow = nRow + 2 + offset2

    def merge_cells(row, col):
        my_ws.merge_cells(start_row=row, start_column=col,
                          end_row=row + nRow, end_column=col)

    def set_border():
        for row in range(nRow + 1):
            cell_border(first_row + row, 1, left_border)
            cell_border(first_row + row, 7, right_border)
            cell_border(first_row + row, 9, right_border)
            cell_border(first_row + row, 13, right_border)
            cell_border(first_row + row, 15, right_border)
            cell_border(first_row + row, 18, right_border)
            cell_border(first_row + row, 20, right_border)
            cell_border(first_row + row, 25, right_border)
            cell_border(first_row + row, 26, right_border)
            cell_border(first_row + row, 27, right_border)

        for col in range(27):
            cell_border(first_row, col + 1, top_border)
            cell_border(first_row + nRow, col + 1, bottom_border)

    merge_cells(first_row, 1)
    merge_cells(first_row, 8)
    merge_cells(first_row, 9)
    merge_cells(first_row, 14)
    merge_cells(first_row, 15)
    merge_cells(first_row, 19)
    merge_cells(first_row, 20)
    merge_cells(first_row, 25)
    merge_cells(first_row, 26)
    merge_cells(first_row, 27)

    set_border()


def mergeCellsPr(first_row, span):
    if span:
        nRow = len(span.linelevels) - 1

    def operation(row, col):
        my_ws.merge_cells(start_row=row, start_column=col, end_row=row + nRow, end_column=col)

    def set_border():
        for row in range(nRow + 1):
            cell_border(first_row + row, 2, right_border)
            cell_border(first_row + row, 21, right_border)
            cell_border(first_row + row, 22, right_border)
            cell_border(first_row + row, 23, right_border)
            cell_border(first_row + row, 24, right_border)

    operation(first_row, 2)
    operation(first_row, 21)
    operation(first_row, 22)
    operation(first_row, 23)
    operation(first_row, 24)

    set_border()


def formatCell(r, c, v):
    my_ws.cell(row=r, column=c, value=v).alignment = alignment
    my_ws.cell(row=r, column=c, height=9).font = font_big


# функции вписывания данных в лист Excel
def write_line_levels(row, line_level):
    set_cell_height(row, 9)
    my_ws.cell(row=row, column=3, value=line_level.wire_mark).alignment = alignment
    my_ws.cell(row=row, column=3).font = font_small
    my_ws.cell(row=row, column=4, value=line_level.wires_amount).alignment = alignment
    my_ws.cell(row=row, column=4).font = font_small
    my_ws.cell(row=row, column=5, value=line_level.diameter).alignment = alignment
    my_ws.cell(row=row, column=5).font = font_small
    my_ws.cell(row=row, column=6, value=line_level.weight).alignment = alignment
    my_ws.cell(row=row, column=6).font = font_small
    my_ws.cell(row=row, column=7, value=line_level.level_height).alignment = alignment
    my_ws.cell(row=row, column=7).font = font_small


def write_spans(row, span):
    my_ws.cell(row=row, column=2, value=span.length).alignment = alignment
    my_ws.cell(row=row, column=2).font = font_big

    def align(index):
        my_ws.cell(row=row + index, column=10, value=span.Gc[index]).alignment = alignment
        my_ws.cell(row=row + index, column=10).font = font_small
        my_ws.cell(row=row + index, column=11, value=span.Pm[index]).alignment = alignment
        my_ws.cell(row=row + index, column=11).font = font_small
        my_ws.cell(row=row + index, column=12, value=span.Gmp[index]).alignment = alignment
        my_ws.cell(row=row + index, column=12).font = font_small
        my_ws.cell(row=row + index, column=13, value=span.Qm[index]).alignment = alignment
        my_ws.cell(row=row + index, column=13).font = font_small
        my_ws.cell(row=row + index, column=16, value=span.M12[index]).alignment = alignment
        my_ws.cell(row=row + index, column=16).font = font_small
        my_ws.cell(row=row + index, column=17, value=span.M13[index]).alignment = alignment
        my_ws.cell(row=row + index, column=17).font = font_small
        my_ws.cell(row=row + index, column=18, value=span.M134[index]).alignment = alignment
        my_ws.cell(row=row + index, column=18).font = font_small

    for index in range(len(span.linelevels)):
        align(index)


def write_supports(row, support):
    my_ws.cell(row=row, column=1, value=support.number).alignment = alignment
    my_ws.cell(row=row, column=1).font = font_big
    my_ws.cell(row=row, column=8, value=support.lampsAmount).alignment = alignment
    my_ws.cell(row=row, column=8).font = font_big
    my_ws.cell(row=row, column=9, value=support.lampHeight).alignment = alignment
    my_ws.cell(row=row, column=9).font = font_big
    my_ws.cell(row=row, column=14, value=support.Wm).alignment = alignment
    my_ws.cell(row=row, column=14).font = font_big
    my_ws.cell(row=row, column=15, value=support.Wg).alignment = alignment
    my_ws.cell(row=row, column=15).font = font_big
    my_ws.cell(row=row, column=19, value=support.MWm).alignment = alignment
    my_ws.cell(row=row, column=19).font = font_big
    my_ws.cell(row=row, column=20, value=support.MWm).alignment = alignment
    my_ws.cell(row=row, column=20).font = font_big
    my_ws.cell(row=row, column=21, value=support.m1[0]).alignment = alignment
    my_ws.cell(row=row, column=21).font = font_big
    my_ws.cell(row=row, column=22, value=support.m2[0]).alignment = alignment
    my_ws.cell(row=row, column=22).font = font_big
    my_ws.cell(row=row, column=23, value=support.m3[0]).alignment = alignment
    my_ws.cell(row=row, column=23).font = font_big
    my_ws.cell(row=row, column=24, value=support.maxMoments[0]).alignment = alignment
    my_ws.cell(row=row, column=24).font = font_big
    if support.span2 is not None:
        dRow = len(support.span1.linelevels)
        my_ws.cell(row=row + dRow, column=21, value=support.m1[1]).alignment = alignment
        my_ws.cell(row=row + dRow, column=21).font = font_big
        my_ws.cell(row=row + dRow, column=22, value=support.m2[1]).alignment = alignment
        my_ws.cell(row=row + dRow, column=22).font = font_big
        my_ws.cell(row=row + dRow, column=23, value=support.m3[1]).alignment = alignment
        my_ws.cell(row=row + dRow, column=23).font = font_big
        my_ws.cell(row=row + dRow, column=24, value=support.maxMoments[1]).alignment = alignment
        my_ws.cell(row=row + dRow, column=24).font = font_big

    my_ws.cell(row=row, column=25, value=support.maxMoment).alignment = alignment
    my_ws.cell(row=row, column=25).font = font_big
    my_ws.cell(row=row, column=26, value=support.supportType).alignment = alignment
    my_ws.cell(row=row, column=26).font = font_big
    my_ws.cell(row=row, column=27, value=support.loadCapacity).alignment = alignment
    my_ws.cell(row=row, column=27).font = font_big


def write_spans2(row, span):
    write_spans(row, span)
    number_of_linelevels = len(span.linelevels)
    y = number_of_linelevels - 1
    while y > 0:
        write_line_levels(row, span.linelevels[y])
        y -= 1
        row += 1
    write_line_levels(row, span.linelevels[0])
    return row


# создаем лист в Excel
title = 'КТП-' + str(ST_number) + ' Л' + str(line_number)
my_ws = workbook.create_sheet("расчёт", 0)
file_name = title + '.xlsx'

# наши стили выравнивания и шрифт ячеек
alignment = Alignment(horizontal="center", vertical="center")
font_big = Font(name='Microsoft Sans Serif', size=8)
font_small = Font(name='Microsoft Sans Serif', size=7)

left_border = Border(left=Side(style='thin', color=Color("969696")))
right_border = Border(right=Side(style='thin', color=Color("969696")))
top_border = Border(top=Side(style='thin', color=Color("969696")))
bottom_border = Border(bottom=Side(style='thin', color=Color("969696")))

# вписываем в ячейки Excel
my_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=27)
my_ws['A1'] = 'КТП-' + str(ST_number) + ' Л' + str(line_number)
my_ws['A1'].font = Font(name='Microsoft Sans Serif', size=11)

# формируем ширину столбцов
col_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
             'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z', 'AA']
col_width = [4.43, 6.86, 7.71, 5.57, 6.86, 7.14, 8.43, 7.14, 7, 7.57, 8.43, 7, 8.71, 8.43, 8.43,
             8.43, 8.43, 8.43,
             8.43, 8.14, 8.43, 8.43, 11.29, 9.43, 8.43, 6.57, 9.43]
table_width = dict(zip(col_names, col_width))
for key, value in table_width.items():
    set_cell_width(key, value + 0.72)

# делаем цикл по опорам
# заполняем таблицу с этого номера
cur_row = 2

for support in supports:
    # делаем расчеты
    for linelevel in support.span1.linelevels:
        support.span1.calculate_parameters(linelevel)
    if support.span2:
        for linelevel in support.span2.linelevels:
            support.span1.calculate_parameters(linelevel)
    support.calculate_parameters()

    # вписываем
    merge_cells_of_support(cur_row, support)
    write_supports(cur_row, support)
    mergeCellsPr(cur_row, support.span1)

    cur_row = write_spans2(cur_row, support.span1)
    if support.span2:
        mergeCellsPr(cur_row + 1, support.span2)
        cur_row = write_spans2(cur_row + 1, support.span2)
    cur_row += 1

for col in range(27):
    cell_border(cur_row, col + 1, top_border)

workbook.save(full_filename)
print('обработан файл', full_filename)
