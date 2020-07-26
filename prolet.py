from yarus import yarus
from math import log
from math import sqrt


class prolet:
    def __init__(self, MPiN, prolety, dlina, svet=False):
        self.MPiN = MPiN
        self.svet = svet
        self.prolety = prolety
        self.dlina = dlina
        self.op1 = self.prolety[0]
        self.op2 = self.prolety[1]
        self.yar = []
        self.yar.append(yarus('ОКСН', h=5.9))
        if self.svet is True:
            self.yar.append(yarus('СІП-2x25', h=6.4))

        # print('пролёт ', self.op1, '-', self.op2)
        # input1 = input('Хотите вручную ввести данные поярусно? [нет/да]: => ')
        # if input1 == 'д' or input1 == 'Д' or input1 == 'да':
        #     input2 = input('Сколько ярусов (без интернет-кабеля и фонарки)? => ')
        #     if int(input2):
        #         for i in range(int(input2)):
        #             text = 'марка провода на ярусе ' + str(i + 1) + ' => '
        #             marka_i = input(text)
        #             h_i = float(input('высота от земли => '))
        #             self.yar.append(yarus(marka_i, h=h_i))
        # else:
        #     self.yar.append(yarus(self.MPiN, h=7.5))

        self.yar.append(yarus(self.MPiN, h=7.5))

        self.Gc = []
        self.Pm = []
        self.Gmp = []
        self.Qm = []
        self.M12 = []
        self.M13 = []
        self.M134 = []

    def calc_prolety(self, yar: object) -> object:
        f1 = round(9.8 * self.dlina * yar.ves * yar.Nprov / 1000, 1)
        f2 = round(500 * 0.9 * 0.85 * 1.2 * yar.diam * (2.6 - 0.3 * log(500 * 0.9)) * 1.5 * (
            1.7 - 0.12 * log(self.dlina)) * self.dlina * yar.Nprov / 1000, 1)
        f3 = round(self.dlina * yar.Nprov * yar.k1, 1)
        f4 = round(self.dlina * yar.Nprov * yar.k2, 1)
        self.M12.append(round(sqrt(f1 ** 2 + f2 ** 2) * yar.h))
        self.M13.append(round((f1 + f3) * yar.h))
        self.M134.append(round(sqrt((f1 + f3) ** 2 + f4 ** 2) * yar.h))
        self.Gc.append(f1)
        self.Pm.append(f2)
        self.Gmp.append(f3)
        self.Qm.append(f4)

        def has_oporu(self, opora):
            if self.op1 == opora or self.op2 == opora:
                return True
