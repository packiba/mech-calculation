import re


def run(fn, wb):
    ws = wb.active
    sup1_sup2 = []
    col_C = ws['C']
    col_H = ws['H']
    col_I = ws['I']

    def formattedList(unformatted_list):
        formatted_list = []
        for item in unformatted_list:
            if item.value:
                formatted_list.append(item)
        return formatted_list

    target_support = formattedList(col_H)
    target_lamp = formattedList(col_I)

    def cellValue(target_row, target_col):
        return ws.cell(row=target_row, column=target_col).value

    def check(sup1sup2, targetCol):

        if sup1sup2 is None:
            return False
        if sup1sup2 == 'Номер пролёта':
            return True
        mFormat = re.findall(r'\d+', sup1sup2)
        flag = False
        if len(mFormat) == 1:
            return False
        if mFormat:
            sup1_sup2.append(mFormat)
            a = 0
            b = 0
            for support in targetCol:
                if support.value is None:
                    continue
                number = support.value
                number1 = int(mFormat[0])
                if number == number1:
                    a = 1
            for support in targetCol:
                if support.value is None:
                    continue
                number = support.value
                number1 = int(mFormat[1])
                if number == number1:
                    b = 1
            if a == 1 and b == 1:
                flag = True
        return flag

    quantity = len(col_C)
    table_old = []
    table_new = []
    for line in range(quantity + 1):
        new_row = []
        for col in range(5):
            new_row.append(cellValue(line + 1, col + 1))
        if check(cellValue(line + 1, 3), target_lamp) is True and \
                cellValue(line + 1, 3) != 'Номер пролёта':
            new_row.append('фонарка')

        table_old.append(new_row)

    title = 'отформатированный'
    my_ws = wb.create_sheet(title)
    my_ws.title = title
    cur_line = 1
    length = 0.0

    for line in table_old:
        span_support_numbers = line[2]
        if check(span_support_numbers, target_support) is True:
            table_new.append(line)
            if line[4] != 'Довжина прольоту, м':
                length += float(line[4])
            print(line)
            number_of_cols = len(line)
            for col in range(number_of_cols):
                my_ws.cell(row=cur_line, column=col + 1, value=line[col])
                # print(cur_line, col + 1, line[col])
            cur_line += 1
    my_ws.cell(row=cur_line, column=5, value=round(length, 1))
    print('Общая длина -', round(length, 1))
    print('Количество опор -', len(target_support))

    wb.save(fn)
    print('записали отформатированную книгу')
