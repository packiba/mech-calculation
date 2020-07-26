import re


def run(fn, wb):
    ws = wb.active
    op1_op2 = []
    col_C = ws['C']
    col_H = ws['H']
    col_I = ws['I']

    def clearList(list):
        newList = []
        for item in list:
            if item.value:
                newList.append(item)
        return newList

    target_op = clearList(col_H)
    target_svet = clearList(col_I)

    def cellValue(row, col):
        return ws.cell(row=row, column=col).value

    def check(o1o2, targetCol):

        if o1o2 is None:
            return False
        if o1o2 == 'Номер прольоту':
            return True
        mFormat = re.findall(r'\d+', o1o2)
        flag = False
        if len(mFormat) == 1:
            return False
        if mFormat:
            op1_op2.append(mFormat)
            a = 0
            b = 0
            for opora in targetCol:
                if opora.value is None:
                    continue
                nomer = opora.value
                nomer1 = int(mFormat[0])
                if nomer == nomer1:
                    a = 1
            for opora in targetCol:
                if opora.value is None:
                    continue
                nomer = opora.value
                nomer1 = int(mFormat[1])
                if nomer == nomer1:
                    b = 1
            if a == 1 and b == 1:
                flag = True
        return flag

    koli4 = len(col_C)
    table_old = []
    table_new = []
    for line in range(koli4 + 1):
        row = []
        for col in range(5):
            row.append(cellValue(line + 1, col + 1))
        if check(cellValue(line + 1, 3), target_svet) is True and cellValue(line + 1, 3) != 'Номер прольоту':
            row.append('фонарка')

        table_old.append(row)

    title = 'отформатированный'
    my_ws = wb.create_sheet(title)
    my_ws.title = title
    l = 1
    dl = 0.0

    for line in table_old:
        o1_o2 = line[2]
        if check(o1_o2, target_op) is True:
            table_new.append(line)
            if line[4] != 'Довжина прольоту, м':
                dl += float(line[4])
            print(line)
            kolCol = len(line)
            for col in range(kolCol):
                my_ws.cell(row=l, column=col + 1, value=line[col])
            l += 1
    my_ws.cell(row=l, column=5, value=round(dl, 1))
    print('Общая длина -', round(dl, 1))
    print('Количество опор -', len(target_op))


    wb.save(fn)
    print('записали отформатированную книгу')
