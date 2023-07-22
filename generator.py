import xlsxwriter as ex
from faker import Faker
import csv
import zipfile
import os

fake = Faker("ru_RU")


class Generator:
    """Генератор файлов с данными."""
    def __init__(self, name: str, format: str = 'xlxs',
                 number_of_strings: int = 100):
        self.name = name
        self.format = format
        self.number_of_strings = number_of_strings
        self.file_name: str = self.name + '.' + self.format
        self.data: list = []
        """Создадим список со строками информации."""
        for row in range(self.number_of_strings):
            new_row = [fake.name(),
                       fake.city(),
                       fake.street_address(),
                       fake.postcode(),
                       fake.job(),
                       fake.phone_number(),
                       fake.hostname(),
                       fake.ascii_free_email(),
                       fake.uri(),
                       fake.company(),
                       fake.city()]
            self.data.append(new_row)

    def run_generator(self):
        """Создает файлы форматов xlsx и csv."""
        if self.format == 'xlsx':
            workbook = ex.Workbook(self.file_name)
            worksheet = workbook.add_worksheet()
            for row in range(self.number_of_strings):
                if (row != 0) and (row % 65530 == 0):
                    worksheet = workbook.add_worksheet()
                col: int = 0
                for data_in_table_cell in self.data[row]:
                    worksheet.write(row - 65530 * (row // 65530), col,
                                    data_in_table_cell)
                    col += 1
            workbook.close()
        elif self.format == 'csv':
            with open(self.file_name, 'w') as f:
                writer = csv.writer(f)
                for row in self.data:
                    writer.writerow(row)
        print(f'Файл {self. file_name} создан')


class Archivator:
    BYTES_IN_KB: int = 1024

    def __init__(self, file_name: str,
                 arc_name: str,
                 size_in_KB: int,
                 arc_format: str = 'zip'):
        self.file_name = file_name
        self.arc_name = arc_name
        self.size_in_KB = size_in_KB
        self.arc_format = arc_format
        self.archieve_name: str = self.arc_name + '.' + self.arc_format

    def make_archieve(self):
        # Cоздаем архив файла
        with zipfile.ZipFile(self.archieve_name, mode='w') as archive:
            archive.write(self.file_name)
        # Разделяем архив на тома
        cur_vol: int = 1  # текущий номер тома
        written: int = 0  # сколько байт записали
        with open(self.archieve_name, 'rb') as src:
            while True:
                output_name = (f'{self.arc_name}{str(cur_vol)}'
                               f'.{self.arc_format}')
                output = open(output_name, 'wb')
                while written < (self.size_in_KB * self.BYTES_IN_KB):
                    data = src.read(self.size_in_KB * self.BYTES_IN_KB)
                    if data == b'':
                        break
                    output.write(data)
                    written += len(data)
                    print('write', len(data), 'bytes to', output_name)
                else:
                    output.close()
                    cur_vol += 1
                    written = 0
                    continue
                output.close()
                break
        # Удаляем исходный архив
        os.remove(self.archieve_name)
        # Формируем один архив
        with zipfile.ZipFile(self.archieve_name, mode='w') as archive:
            for cur in range(1, cur_vol+1):
                cur_name = (f'{self.arc_name}{str(cur)}'
                            f'.{self.arc_format}')
                archive.write(cur_name)
                os.remove(cur_name)
        print('Процесс архивирования завершен.')


def main():
    print('Привет, давай сгенерируем файл с данными!')
    name_file: str = input('Введи имя генерируемого файла: ')
    if name_file == '':
        raise ValueError('Необходимо ввести название файла.')
    file_formats = {'xlsx': 'xlsx', 'csv': 'csv'}
    format_file: str = input('Введи формат генерируемого файла: ')
    try:
        format_file = file_formats[format_file]
    except KeyError:
        raise KeyError('Такого формата нет в базе')
    number_of_strings = input('Введи количество строк данных: ')
    try:
        int(number_of_strings)
    except ValueError:
        raise ValueError('Необходимо ввести целое число')
    Generator(name_file, format_file, int(number_of_strings)).run_generator()
    print('Теперь создадим архив нашего файла!')
    arc_name = input('Введи имя архива или оставь это поле пустым: ')
    if arc_name == '':
        arc_name = f'archive_of_{name_file}'
    arc_format = input('Введи формат архива: ')
    arc_formats = {'zip': 'zip', '7z': '7z'}
    try:
        arc_format = arc_formats[arc_format]
    except KeyError:
        raise KeyError('Такого формата нет в базе')
    size = input('Введи предельный размер тома архива: ')
    try:
        int(size)
    except ValueError:
        raise ValueError('Необходимо ввести целое число')
    Archivator(name_file + '.' + format_file, arc_name,
               int(size),
               arc_format).make_archieve()


main()
print('done')
