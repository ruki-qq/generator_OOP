import xlsxwriter as ex
from faker import Faker
import csv
import zipfile
import os

fake = Faker("ru_RU")


class Generator:
    """Генератор файлов с данными с функцией упаковщика."""
    BYTES_IN_KB: int = 1024

    def __init__(self):
        self.file_formats = {'xlsx': 'xlsx', 'csv': 'csv'}
        self.arc_formats = {'zip': 'zip', '7z': '7z'}

    def run_generator(self):
        """Создает файлы форматов xlsx и csv."""
        """Просим пользователя ввести необходимые данные и проверяем их."""
        print('Привет, давай сгенерируем файл с данными!')
        self.name_file: str = input('Введи имя генерируемого файла: ')
        if self.name_file == '':
            raise ValueError('Необходимо ввести название файла.')
        self.format_file: str = input('Введи формат генерируемого файла: ')
        try:
            self.format_file = self.file_formats[self.format_file]
        except KeyError:
            raise KeyError('Такого формата нет в базе')
        self.number_of_strings = input('Введи количество строк данных: ')
        try:
            int(self.number_of_strings)
        except ValueError:
            raise ValueError('Необходимо ввести целое число')
        self.full_file_name: str = self.name_file + '.' + self.format_file
        # Создадим список с фейковыми данными.
        self.data: list = []
        for row in range(int(self.number_of_strings)):
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
        # Создаем файл с данными из списка.
        if self.format_file == 'xlsx':
            workbook = ex.Workbook(self.full_file_name)
            worksheet = workbook.add_worksheet()
            for row in range(int(self.number_of_strings)):
                if (row != 0) and (row % 65530 == 0):
                    worksheet = workbook.add_worksheet()
                col: int = 0
                for data_in_table_cell in self.data[row]:
                    worksheet.write(row - 65530 * (row // 65530), col,
                                    data_in_table_cell)
                    col += 1
            workbook.close()
        elif self.format_file == 'csv':
            with open(self.full_file_name, 'w') as f:
                writer = csv.writer(f)
                for row in self.data:
                    writer.writerow(row)
        print(f'Файл {self.full_file_name} создан')

    def make_archieve(self):
        # Создает архив из файла.
        # Просим пользователя ввести данные и проверяем их.
        print('Теперь создадим архив нашего файла!')
        self.arc_name = input('Введи имя архива или оставь это поле пустым: ')
        if self.arc_name == '':
            self.arc_name = f'archive_of_{self.name_file}'
        self.arc_format = input('Введи формат архива: ')
        try:
            self.arc_format = self.arc_formats[self.arc_format]
        except KeyError:
            raise KeyError('Такого формата нет в базе')
        self.archieve_name: str = self.arc_name + '.' + self.arc_format
        self.archieve_size = input('Введи предельный размер тома архива: ')
        try:
            int(self.archieve_size)
        except ValueError:
            raise ValueError('Необходимо ввести целое число')
        # Cоздаем архив файла
        with zipfile.ZipFile(self.archieve_name, mode='w') as archive:
            archive.write(self.full_file_name)
        # Разделяем архив на тома
        cur_vol: int = 1  # текущий номер тома
        written: int = 0  # сколько байт записали
        with open(self.archieve_name, 'rb') as src:
            while True:
                output_name = (f'{self.arc_name}{str(cur_vol)}'
                               f'.{self.arc_format}')
                output = open(output_name, 'wb')
                while written < (int(self.archieve_size) * self.BYTES_IN_KB):
                    data = src.read(int(self.archieve_size) * self.BYTES_IN_KB)
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

    def add_archieve_format(self, format: str):
        # Добавляет новый формат архива.
        self.arc_formats[format] = format


def main():
    A = Generator()
    A.add_archieve_format('gzip')
    A.run_generator()
    A.make_archieve()


main()
print('done')
