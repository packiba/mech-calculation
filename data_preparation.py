import re


def run(filename, workbook):
    worksheet = workbook.active
    col_C = worksheet['C']  # опоры пролётов '1 - 2'
    col_H = worksheet['H']  # опоры для расчёта (на которых подвес ОКСН)
    col_I = worksheet['I']  # опоры, на которых фонарка

    supports_for_calc = list(item.value for item in col_H if item.value)  # очищаем от пустых ячеек
    supports_with_lamp = list(item.value for item in col_I if item.value)  # очищаем от пустых ячеек

    def cellValue(target_row, target_col):
        return worksheet.cell(row=target_row, column=target_col).value

    def supports_list(span): # span = '33 - 45' => ['33', '45']
        return re.findall(r'\d+', span)

    def is_span_included_in_list(span, support_list):  # span = '33 - 45'
        span_support_numbers = supports_list(span) # ['33', '45']
        if len(span_support_numbers) == 0:
            return False
        if len(span_support_numbers) == 1:
            number_of_sup = int(span_support_numbers[0])
            if number_of_sup in support_list:
                return True
        if len(span_support_numbers) == 2:
            number_of_sup1 = int(span_support_numbers[0])
            number_of_sup2 = int(span_support_numbers[1])
            if number_of_sup1 in support_list and number_of_sup2 in support_list:
                return True
        return False

    total_number_of_spans = len(col_C)
    table_old = []
    for line_number in range(total_number_of_spans):
        new_row = []
        for col_number in range(5):
            new_row.append(cellValue(line_number + 1, col_number + 1))
        span = new_row[2]
        if span != 'Номер прольоту':
            if is_span_included_in_list(span, supports_with_lamp):
                new_row.append('есть')
        else:
            new_row.append('фонарка')
        table_old.append(new_row)


    table_new = []
    length = 0.0
    for line in table_old:
        span = line[2]
        if span == 'Номер прольоту':
            table_new.append(line)
            print(line)
            continue
        if is_span_included_in_list(span, supports_for_calc):
            table_new.append(line)
            length += float(line[4])
            print(line)

    # title = 'отформатированный'
    new_worksheet = workbook.create_sheet('отформатированный', 0)
    for line_number in range(len(table_new)):
        for col_number in range(len(table_new[line_number])):
            new_worksheet.cell(row=line_number + 1, column=col_number + 1, value=table_new[line_number][col_number])

    new_worksheet.cell(row=line_number + 2, column=5, value=round(length, 1))
    print('Общая длина -', round(length, 1))
    print('Количество опор -', len(supports_for_calc))

    workbook.save(filename)
    print('записали отформатированную книгу')
