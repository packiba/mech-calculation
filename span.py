from linelevel import LineLevel
from math import log
from math import sqrt
import numpy as np


class Span:
    """Пролёт между опорами"""

    def __init__(self, mark, span_support_numbers, length, lamp=False):
        self.mark = mark
        self.span_support1_number = span_support_numbers[0]
        self.span_support2_number = span_support_numbers[1]
        self.length = length
        self.lamp = lamp
        self.M12 = np.array([], float)
        self.M13 = np.array([], float)
        self.M134 = np.array([], float)
        self.Gc = np.array([], float)
        self.Pm = np.array([], float)
        self.Gmp = np.array([], float)
        self.Qm = np.array([], float)
        self.linelevels = [] # список объектов ярусов

        self.linelevels.append(LineLevel('ОКСН', 5.9))
        if self.lamp:
            self.linelevels.append(LineLevel('СІП-2x25', 6.4))
        self.linelevels.append(LineLevel(self.mark, 7.5))


    def calculate_parameters(self, linelevel):
        f1 = round(9.8 * self.length * linelevel.weight * linelevel.number_of_wires / 1000, 1)
        f2 = round(500 * 0.9 * 0.85 * 1.2 * linelevel.diameter * (2.6 - 0.3 * log(500 * 0.9)) * 1.5
                   * (1.7 - 0.12 * log(self.length)) * self.length * linelevel.number_of_wires / 1000,
                   1)
        f3 = round(self.length * linelevel.number_of_wires * linelevel.k1, 1)
        f4 = round(self.length * linelevel.number_of_wires * linelevel.k2, 1)

        self.M12 = np.append(self.M12, round(sqrt(f1 ** 2 + f2 ** 2) * linelevel.level_height))
        self.M13 = np.append(self.M13, round((f1 + f3) * linelevel.level_height))
        self.M134 = np.append(self.M134, round(sqrt((f1 + f3) ** 2 + f4 ** 2) * linelevel.level_height))
        self.Gc = np.append(self.Gc, f1)
        self.Pm = np.append(self.Pm, f2)
        self.Gmp = np.append(self.Gmp, f3)
        self.Qm = np.append(self.Qm, f4)
