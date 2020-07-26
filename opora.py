class opora:
    def __init__(self, nomer, prolet1, prolet2=None):
        self.nomer = nomer
        self.prolet1 = prolet1
        self.prolet2 = prolet2
        self.Nsvet = 0
        self.hsvet = 0
        self.Wm = 826
        self.Wg = 485
        self.MWm = 3098
        self.MWg = 1818
        self.maxMoments = []
        self.tipStoiki = 'СВ95-2'
        self.nesSposobnOp = 29420
        self.m1 = []
        self.m2 = []
        self.m3 = []
        self.maxMomOp = 0
        if self.prolet1.svet is True:
            self.Nsvet = 1
            self.hsvet = 6.4
        if self.prolet2 and self.prolet2.svet is True:
            self.Nsvet = 1
            self.hsvet = 6.4

    def calc(self):
        for p in [self.prolet1, self.prolet2]:
            sumM1 = 0
            sumM2 = 0
            sumM3 = 0
            if p is None:
                break
            for j in p.yar:
                p.calc_prolety(j)
                sumM1 = sumM1 + p.M12[-1]
                sumM2 = sumM2 + p.M13[-1]
                sumM3 = sumM3 + p.M134[-1]
            m1 = sumM1 + self.MWm
            m2 = round(sumM2)
            m3 = round(0.9 * (sumM3 + self.MWg))
            sumMoments_row = [m1, m2, m3]
            maxMpr = max(sumMoments_row)
            self.maxMoments.append(maxMpr)
            self.m1.append(m1)
            self.m2.append(m2)
            self.m3.append(m3)
        if (len(self.maxMoments)) < 2:
            self.maxMomOp = self.maxMoments[0]
        else:
            self.maxMomOp = max(self.maxMoments[0], self.maxMoments[1])
        if self.maxMomOp > self.nesSposobnOp:
            print('не проходит опора №', self.nomer)