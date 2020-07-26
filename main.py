from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Color
from openpyxl.styles.borders import Border, Side
import re
from prepareData import run

def makeListOpor(listProletov: list):
    k = 0
    listOpor = []
    listKoncIndex = []
    for x, y in listProletov:
        if k != 0 and k != x:
            listKoncIndex.append(int(k))
        if x != k:
            listOpor.append(int(x))
            listKoncIndex.append(int(x))
        listOpor.append(int(y))
        k = y
    listKoncIndex.append(int(k))
    return (listOpor, listKoncIndex)

def flag_konc(op, list):
    for item in list:
        if item == op:
            return True
    return False

def go(fn):
    from prolet import prolet
    from opora import opora
    # открываем файл и берем текущий лист
    # fn1 = input('Enter file name: ')
    # fn = fn1 + '.xlsx'
    wb = load_workbook(fn)
    # подготавливаем список из файла
    run(fn, wb)
    # загружаем отформатированный список Excel
    wb = load_workbook(fn)
    ws = wb['отформатированный']
    # номер КТП
    val_A2 = ws['A2'].value
    Ntp = int(re.findall(r'\d+', val_A2)[0])
    # номер линии
    val_B2 = ws['B2'].value
    Nlin = int(re.findall(r'\d+', val_B2)[0])
    # список пролётов
    prolety = []
    col_C = ws['C']
    for item in col_C:
        if item.value is None:
            break
        mFormat = re.findall(r'\d+', item.value)
        if mFormat:
            prolety.append(mFormat)

    # список опор (вызываем функцию makeListOpor)
    opory = makeListOpor(prolety)[0]
    koncOpory = makeListOpor(prolety)[1]

    kolOpor = len(makeListOpor(prolety)[0])
    kolProletov = len(prolety)

    marki_kol = []
    dliny = []
    svet = []
    for i in range(0, kolProletov):
        d = ws.cell(row=i + 2, column=4)
        e = ws.cell(row=i + 2, column=5)
        f = ws.cell(row=i + 2, column=6)
        marki_kol.append(d.value)
        dliny.append(round(e.value))
        svet.append(f.value)

    # создаем список с объектами пролетов
    proList = []
    for mark, prol, dlin, sv in zip(marki_kol, prolety, dliny, svet):
        flag = False
        if sv == 'фонарка':
            flag = True
        p = prolet(mark, prol, dlin, flag)
        proList.append(p)

    # создаём список с объектами опор
    oporList = []
    vrem = None
    for my_op in opory:
        for prolet in proList:
            if flag_konc(my_op, koncOpory):
                if int(prolet.op1) == my_op or int(prolet.op2) == my_op:
                    op = opora(my_op, prolet)
                    oporList.append(op)
                    break
            else:
                if int(prolet.op2) == my_op:
                    vrem = prolet
                    continue
                elif int(prolet.op1) == my_op:
                    op = opora(my_op, vrem, prolet)
                    oporList.append(op)
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


    def mergeCellsOp(first_row, opora):
        delt1 = 0
        delt2 = 0
        if opora.prolet1.svet is True:
            delt1 = 1
        nRow = 1 + delt1
        if opora.prolet2:
            if opora.prolet2.svet is True:
                delt2 = 1
            nRow = nRow + 2 + delt2

        def operation(row, col):
            my_ws.merge_cells(start_row=row, start_column=col,
                              end_row=row + nRow, end_column=col)

        def set_border():
            for row in range(nRow + 1):
                cell_border(first_row + row, 1, leftB)
                cell_border(first_row + row, 7, rightB)
                cell_border(first_row + row, 9, rightB)
                cell_border(first_row + row, 13, rightB)
                cell_border(first_row + row, 15, rightB)
                cell_border(first_row + row, 18, rightB)
                cell_border(first_row + row, 20, rightB)
                cell_border(first_row + row, 25, rightB)
                cell_border(first_row + row, 26, rightB)
                cell_border(first_row + row, 27, rightB)

            for col in range(27):
                cell_border(first_row, col + 1, topB)
                cell_border(first_row + nRow, col + 1, bottomB)

        operation(first_row, 1)
        operation(first_row, 8)
        operation(first_row, 9)
        operation(first_row, 14)
        operation(first_row, 15)
        operation(first_row, 19)
        operation(first_row, 20)
        operation(first_row, 25)
        operation(first_row, 26)
        operation(first_row, 27)

        set_border()


    def mergeCellsPr(first_row, prolet):
        if prolet:
            nRow = len(prolet.yar) - 1

        def operation(row, col):
            my_ws.merge_cells(start_row=row, start_column=col, end_row=row + nRow, end_column=col)

        def set_border():
            for row in range(nRow + 1):
                cell_border(first_row + row, 2, rightB)
                cell_border(first_row + row, 21, rightB)
                cell_border(first_row + row, 22, rightB)
                cell_border(first_row + row, 23, rightB)
                cell_border(first_row + row, 24, rightB)

        operation(first_row, 2)
        operation(first_row, 21)
        operation(first_row, 22)
        operation(first_row, 23)
        operation(first_row, 24)

        set_border()


    def formatCell(r, c, v):
        my_ws.cell(row=r, column=c, value=v).alignment = al
        my_ws.cell(row=r, column=c, height=9).font = f


    # функции вписывания данных в лист Excel
    def write_yarusy(row, yar):
        set_cell_height(row, 9)
        my_ws.cell(row=row, column=3, value=yar.marka).alignment = al
        my_ws.cell(row=row, column=3).font = f2
        my_ws.cell(row=row, column=4, value=yar.Nprov).alignment = al
        my_ws.cell(row=row, column=4).font = f2
        my_ws.cell(row=row, column=5, value=yar.diam).alignment = al
        my_ws.cell(row=row, column=5).font = f2
        my_ws.cell(row=row, column=6, value=yar.ves).alignment = al
        my_ws.cell(row=row, column=6).font = f2
        my_ws.cell(row=row, column=7, value=yar.h).alignment = al
        my_ws.cell(row=row, column=7).font = f2


    def write_prolety(row, prolet):
        my_ws.cell(row=row, column=2, value=prolet.dlina).alignment = al
        my_ws.cell(row=row, column=2).font = f

        def oper(ind):
            my_ws.cell(row=row + ind, column=10, value=prolet.Gc[ind]).alignment = al
            my_ws.cell(row=row + ind, column=10).font = f2
            my_ws.cell(row=row + ind, column=11, value=prolet.Pm[ind]).alignment = al
            my_ws.cell(row=row + ind, column=11).font = f2
            my_ws.cell(row=row + ind, column=12, value=prolet.Gmp[ind]).alignment = al
            my_ws.cell(row=row + ind, column=12).font = f2
            my_ws.cell(row=row + ind, column=13, value=prolet.Qm[ind]).alignment = al
            my_ws.cell(row=row + ind, column=13).font = f2
            my_ws.cell(row=row + ind, column=16, value=prolet.M12[ind]).alignment = al
            my_ws.cell(row=row + ind, column=16).font = f2
            my_ws.cell(row=row + ind, column=17, value=prolet.M13[ind]).alignment = al
            my_ws.cell(row=row + ind, column=17).font = f2
            my_ws.cell(row=row + ind, column=18, value=prolet.M134[ind]).alignment = al
            my_ws.cell(row=row + ind, column=18).font = f2

        for i in range(len(prolet.yar)):
            oper(i)


    def write_opory(row, opora):
        my_ws.cell(row=row, column=1, value=opora.nomer).alignment = al
        my_ws.cell(row=row, column=1).font = f
        my_ws.cell(row=row, column=8, value=opora.Nsvet).alignment = al
        my_ws.cell(row=row, column=8).font = f
        my_ws.cell(row=row, column=9, value=opora.hsvet).alignment = al
        my_ws.cell(row=row, column=9).font = f
        my_ws.cell(row=row, column=14, value=opora.Wm).alignment = al
        my_ws.cell(row=row, column=14).font = f
        my_ws.cell(row=row, column=15, value=opora.Wg).alignment = al
        my_ws.cell(row=row, column=15).font = f
        my_ws.cell(row=row, column=19, value=opora.MWm).alignment = al
        my_ws.cell(row=row, column=19).font = f
        my_ws.cell(row=row, column=20, value=opora.MWm).alignment = al
        my_ws.cell(row=row, column=20).font = f
        my_ws.cell(row=row, column=21, value=opora.m1[0]).alignment = al
        my_ws.cell(row=row, column=21).font = f
        my_ws.cell(row=row, column=22, value=opora.m2[0]).alignment = al
        my_ws.cell(row=row, column=22).font = f
        my_ws.cell(row=row, column=23, value=opora.m3[0]).alignment = al
        my_ws.cell(row=row, column=23).font = f
        my_ws.cell(row=row, column=24, value=opora.maxMoments[0]).alignment = al
        my_ws.cell(row=row, column=24).font = f
        if opora.prolet2 is not None:
            dRow = len(opora.prolet1.yar)
            my_ws.cell(row=row + dRow, column=21, value=opora.m1[1]).alignment = al
            my_ws.cell(row=row + dRow, column=21).font = f
            my_ws.cell(row=row + dRow, column=22, value=opora.m2[1]).alignment = al
            my_ws.cell(row=row + dRow, column=22).font = f
            my_ws.cell(row=row + dRow, column=23, value=opora.m3[1]).alignment = al
            my_ws.cell(row=row + dRow, column=23).font = f
            my_ws.cell(row=row + dRow, column=24, value=opora.maxMoments[1]).alignment = al
            my_ws.cell(row=row + dRow, column=24).font = f

        my_ws.cell(row=row, column=25, value=opora.maxMomOp).alignment = al
        my_ws.cell(row=row, column=25).font = f
        my_ws.cell(row=row, column=26, value=opora.tipStoiki).alignment = al
        my_ws.cell(row=row, column=26).font = f
        my_ws.cell(row=row, column=27, value=opora.nesSposobnOp).alignment = al
        my_ws.cell(row=row, column=27).font = f


    def write_prolety2(row, prolet):
        write_prolety(row, prolet)
        Nyar = len(prolet.yar)
        y = Nyar - 1
        while y > 0:
            write_yarusy(row, prolet.yar[y])
            y -= 1
            row += 1
        write_yarusy(row, prolet.yar[0])
        return row


    # создаем лист в Excel
    title = 'КТП-' + str(Ntp) + ' Л' + str(Nlin)
    my_ws = wb.create_sheet("расчёт", 0)
    file_name = title + '.xlsx'

    # наши стили выравнивания и шрифт ячеек
    al = Alignment(horizontal="center", vertical="center")
    f = Font(name='Microsoft Sans Serif', size=8)
    f2 = Font(name='Microsoft Sans Serif', size=7)

    leftB = Border(left=Side(style='thin', color=Color("969696")))
    rightB = Border(right=Side(style='thin', color=Color("969696")))
    topB = Border(top=Side(style='thin', color=Color("969696")))
    bottomB = Border(bottom=Side(style='thin', color=Color("969696")))

    # вписываем в ячейки Excel
    my_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=27)
    my_ws['A1'] = 'КТП-' + str(Ntp) + ' Л' + str(Nlin)
    my_ws['A1'].font = Font(name='Microsoft Sans Serif', size=11)

    # формируем ширину столбцов
    col_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
                 'V', 'W', 'X', 'Y', 'Z', 'AA']
    col_width = [4.43, 6.86, 7.71, 5.57, 6.86, 7.14, 8.43, 7.14, 7, 7.57, 8.43, 7, 8.71, 8.43, 8.43, 8.43, 8.43, 8.43,
                 8.43, 8.14, 8.43, 8.43, 11.29, 9.43, 8.43, 6.57, 9.43]
    table_width = dict(zip(col_names, col_width))
    for key, value in table_width.items():
        set_cell_width(key, value + 0.72)

    # делаем цикл по опорам
    # заполняем таблицу с этого номера
    cur_row = 2

    for opora in oporList:
        # делаем расчеты
        for y in opora.prolet1.yar:
            opora.prolet1.calc_prolety(y)
        if opora.prolet2:
            for y in opora.prolet2.yar:
                opora.prolet1.calc_prolety(y)
        opora.calc()

        # вписываем
        mergeCellsOp(cur_row, opora)
        write_opory(cur_row, opora)
        mergeCellsPr(cur_row, opora.prolet1)

        cur_row = write_prolety2(cur_row, opora.prolet1)
        if opora.prolet2:
            mergeCellsPr(cur_row + 1, opora.prolet2)
            cur_row = write_prolety2(cur_row + 1, opora.prolet2)
        cur_row += 1

    for col in range(27):
        cell_border(cur_row, col + 1, topB)

        # opora = oporList[2]
        # for y in opora.prolet1.yar:
        #     opora.prolet1.calc_prolety(y)
        #     if opora.prolet2 != None:
        #         for y in opora.prolet2.yar:
        #             opora.prolet1.calc_prolety(y)
        #     opora.calc()
        #
        #  # вписываем
        # mergeCellsOp(cur_row, opora)
        # write_opory(cur_row, opora)
        # mergeCellsPr(cur_row, opora.prolet1)
        # cur_row = write_prolety2(cur_row, opora.prolet1)
        # if opora.prolet2 != None:
        #     mergeCellsPr(cur_row + 1, opora.prolet2)
        #     cur_row = write_prolety2(cur_row + 1, opora.prolet2)

    wb.save(fn)
    print('обработан файл', fn)