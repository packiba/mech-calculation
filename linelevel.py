import re


class LineLevel:
    """Ярус ЛЭП с параметрами проводов. MPiN - запись типа А-2x16 (марка провода, количество проводов и их сечение)"""

    reference_data = {
        'А-16': [5.1, 43, 6.2, 3.15],
        'А-25': [6.4, 68, 6.31, 3.2],
        'А-35': [7.5, 94, 6.46, 3.28],
        'А-50': [9, 135, 6.66, 3.38],
        'СІП-2x25': [17.5, 215, 6.69, 3.48],
        'СІП-4x25': [21.1, 403, 7.48, 4.1],
        'СІП-4x35': [24, 511, 7.54, 4.22],
        'СІП-2x16': [12.2, 91, 6.4, 3.29],
        'СІП-4x16': [14.7, 183, 6.5, 3.35],
        'СІП-4x50': [29, 711, 7.62, 4.45],
        'СІП-4x70': [32, 983, 7.74, 4.58],
        'СІП-4x95': [37, 1334, 8.66, 5.51],
        'ОКСН': [8.4, 40, 5.84, 2.75]
    }

    def __init__(self, MPiN, level_height=7.5):  # MPiN='А-2x16'
        self.level_height = level_height
        self.MPiN = MPiN
        self.diameter = 0
        self.weight = 0
        self.k1 = 0
        self.k2 = 0

        if self.MPiN[0] == 'А':
            decomposition1 = re.findall(r'\w+', self.MPiN)  # ['А', '2x16']
            decomposition2 = re.findall(r'\d+', decomposition1[1])  # ['2', '16']
            self.number_of_wires = int(decomposition2[0])
            self.wire_mark = decomposition1[0] + '-' + decomposition2[1]
        else:
            self.number_of_wires = 1
            self.wire_mark = self.MPiN

        self.calculate_parameters(self.wire_mark)

    def calculate_parameters(self, wire_mark):
        if wire_mark in self.reference_data.keys():
            self.diameter = self.reference_data[wire_mark][0]
            self.weight = self.reference_data[wire_mark][1]
            self.k1 = self.reference_data[wire_mark][2]
            self.k2 = self.reference_data[wire_mark][3]
        else:
            print('не могу определить марку провода', wire_mark)
        if wire_mark == 'ОКСН':
            self.level_height = 5.9
