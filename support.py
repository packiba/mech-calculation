class Support:
    """Опора ЛЭП со всеми свойствами"""

    Wm = 826
    Wg = 485
    MWm = 3098
    MWg = 1818
    supportType = 'СВ95-2'
    loadCapacity = 29420

    def __init__(self, number, span1, span2=None):
        self.number = number
        self.span1 = span1
        self.span2 = span2
        self.lampsAmount = 0
        self.lampHeight = 0
        self.maxMoments = []
        self.m1 = []
        self.m2 = []
        self.m3 = []
        self.maxMoment = 0
        if self.span1.lamp is True:
            self.lampsAmount = 1
            self.lampHeight = 6.4
        if self.span2 and self.span2.lamp is True:
            self.lampsAmount = 1
            self.lampHeight = 6.4

    def calculate_parameters(self):
        for p in [self.span1, self.span2]:
            sumM1 = 0
            sumM2 = 0
            sumM3 = 0
            if p is None:
                break
            for j in p.linelevels:
                p.calculate_parameters(j)
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
            self.maxMoment = self.maxMoments[0]
        else:
            self.maxMoment = max(self.maxMoments[0], self.maxMoments[1])
        if self.maxMoment > self.loadCapacity:
            print('не проходит опора №', self.number)
