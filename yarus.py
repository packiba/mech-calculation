import re


class yarus:
    def __init__(self, MPiN, h=7.5):
        self.MPiN = MPiN
        self.h = h
        self.nomer = 0
        self.ves = 0
        self.diam = 0
        self.k1 = 0
        self.k2 = 0

        result = re.findall(r'\w+', self.MPiN)
        if result[0] == 'А':
            result2 = re.findall(r'\d+', result[1])

            self.Nprov = int(result2[0])
            self.marka = result[0] + '-' + result2[1]
        else:
            self.Nprov = 1
            self.marka = self.MPiN

        if self.marka == 'А-16':
            self.diam = 5.1
            self.ves = 43
            self.k1 = 6.2
            self.k2 = 3.15

        if self.marka == 'А-25':
            self.diam = 6.4
            self.ves = 68
            self.k1 = 6.31
            self.k2 = 3.2

        if self.marka == 'А-35':
            self.diam = 7.5
            self.ves = 94
            self.k1 = 6.46
            self.k2 = 3.28

        if self.marka == 'А-50':
            self.diam = 9
            self.ves = 135
            self.k1 = 6.66
            self.k2 = 3.38

        if self.marka == 'СІП-2x25':
            self.diam = 17.5
            self.ves = 215
            self.k1 = 6.69
            self.k2 = 3.48

        if self.marka == 'СІП-4x25':
            self.diam = 21.1
            self.ves = 403
            self.k1 = 7.48
            self.k2 = 4.1

        if self.marka == 'СІП-4x35':
            self.diam = 24
            self.ves = 511
            self.k1 = 7.54
            self.k2 = 4.22

        if self.marka == 'СІП-2x16':
            self.diam = 12.2
            self.ves = 91
            self.k1 = 6.4
            self.k2 = 3.29

        if self.marka == 'СІП-4x16':
            self.diam = 14.7
            self.ves = 183
            self.k1 = 6.5
            self.k2 = 3.35

        if self.marka == 'СІП-4x50':
            self.diam = 29
            self.ves = 711
            self.k1 = 7.62
            self.k2 = 4.45

        if self.marka == 'СІП-4x70':
            self.diam = 32
            self.ves = 983
            self.k1 = 7.74
            self.k2 = 4.58

        if self.marka == 'СІП-4x95':
            self.diam = 37
            self.ves = 1334
            self.k1 = 8.66
            self.k2 = 5.51

        if self.marka == 'ОКСН':
            self.diam = 8.4
            self.ves = 40
            self.k1 = 5.84
            self.k2 = 2.75
            self.h = 5.9
            self.nomer = 1

        if self.k1 == 0:
            print('не могу определить марку провода', self.MPiN)
