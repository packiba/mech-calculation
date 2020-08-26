import os

from initial_data import InitialData
from calculation_data import CalculationData


# dir_path = r"e:\Coding\PycharmProjects\Мехрасчет опор"
dir_path = r"f:\Рабочие проекты\Лёха\Видеонаблюдение с.Баштановка"
files = os.listdir(dir_path)
os.chdir(dir_path)
filenames = list(os.path.abspath(file) for file in files if os.path.splitext(file)[1] == '.xlsx')
total_length = 0
for filename in filenames:
    data = InitialData(filename)
    total_length += data.length
    cal_data = CalculationData(data)
    print('\n',filename)
    print(f'Длина линии - {round(data.length)} м')
print(f'\nОбщая длина по всем линиям - {round(total_length)} м')
