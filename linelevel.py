import re


class LineLevel:
    """Ярус ЛЭП с параметрами проводов"""

    def __init__(self, MPiN, level_height=7.5):
        self.MPiN = MPiN
        self.number = 0
        self.weight = 0
        self.diameter = 0
        self.level_height = level_height
        self.k1 = 0
        self.k2 = 0
        self.wires_amount = 0
        self.wire_mark = ''

        result = re.findall(r'\w+', self.MPiN)
        if result[0] == 'А':
            result2 = re.findall(r'\d+', result[1])
            self.wires_amount = int(result2[0])
            self.wire_mark = result[0] + '-' + result2[1]
        else:
            self.wires_amount = 1
            self.wire_mark = self.MPiN

        if self.wire_mark == 'А-16':
            self.diameter = 5.1
            self.weight = 43
            self.k1 = 6.2
            self.k2 = 3.15

        if self.wire_mark == 'А-25':
            self.diameter = 6.4
            self.weight = 68
            self.k1 = 6.31
            self.k2 = 3.2

        if self.wire_mark == 'А-35':
            self.diameter = 7.5
            self.weight = 94
            self.k1 = 6.46
            self.k2 = 3.28

        if self.wire_mark == 'А-50':
            self.diameter = 9
            self.weight = 135
            self.k1 = 6.66
            self.k2 = 3.38

        if self.wire_mark == 'СІП-2x25':
            self.diameter = 17.5
            self.weight = 215
            self.k1 = 6.69
            self.k2 = 3.48

        if self.wire_mark == 'СІП-4x25':
            self.diameter = 21.1
            self.weight = 403
            self.k1 = 7.48
            self.k2 = 4.1

        if self.wire_mark == 'СІП-4x35':
            self.diameter = 24
            self.weight = 511
            self.k1 = 7.54
            self.k2 = 4.22

        if self.wire_mark == 'СІП-2x16':
            self.diameter = 12.2
            self.weight = 91
            self.k1 = 6.4
            self.k2 = 3.29

        if self.wire_mark == 'СІП-4x16':
            self.diameter = 14.7
            self.weight = 183
            self.k1 = 6.5
            self.k2 = 3.35

        if self.wire_mark == 'СІП-4x50':
            self.diameter = 29
            self.weight = 711
            self.k1 = 7.62
            self.k2 = 4.45

        if self.wire_mark == 'СІП-4x70':
            self.diameter = 32
            self.weight = 983
            self.k1 = 7.74
            self.k2 = 4.58

        if self.wire_mark == 'СІП-4x95':
            self.diameter = 37
            self.weight = 1334
            self.k1 = 8.66
            self.k2 = 5.51

        if self.wire_mark == 'ОКСН':
            self.diameter = 8.4
            self.weight = 40
            self.k1 = 5.84
            self.k2 = 2.75
            self.level_height = 5.9
            self.number = 1

        if self.k1 == 0:
            print('не могу определить марку провода', self.MPiN)
