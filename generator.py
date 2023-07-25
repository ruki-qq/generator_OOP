import xlsxwriter as ex #непонятное название переменной
from faker import Faker
import csv
import zipfile
import os

# ↑↑↑ Есть сторонние либы, но нет requirements.txt, 
# чтобы сразу все поставить, плюс несоблюден порядок импортов

fake = Faker("ru_RU")

# ↑↑↑ Эта переменная будет использоваться только в классе генератора данных,
# логичнее положить ее туда, в метод генерации, чтобы не занимала память, пока не генерим
# плюс нет аннотации типа


class Generator:
    
    # ↑↑↑ Название класса недостаточно информативно, лучше указать, 
    # что именно мы генерируем, например FakeDataGenerator полностью отразить суть
    
    """Генератор файлов с данными с функцией упаковщика."""
    BYTES_IN_KB: int = 1024

    def __init__(self):
        self.file_formats = {'xlsx': 'xlsx', 'csv': 'csv'}
        self.arc_formats = {'zip': 'zip', '7z': '7z'}
        
        # ↑↑↑ Нужно использовать списки и проверять наличие 
        # через 'if elem in list:', плюс нет аннотации типов

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
        # ↑↑↑ Нет аннотации типа
        try:
            int(self.number_of_strings)
        except ValueError:
            raise ValueError('Необходимо ввести целое число')
        
        # ↑↑↑ UI лучше валидировать, заставляя юзера вводить значение, пока оно нас
        # не удовлетворит, например через {while True} и {break}

        # ↑↑↑ Не нужно записывать в self значения, которые мы будем использовать
        # только в одном методе, после того как метод отработал, значения остаются в
        # экземпляре класса и тратят память впустую
        
        # ↑↑↑ Все, что выше в данном методе, относится к user интерфейсу, этот код 
        # должен быть вынесен в отдельный класс в отдельном файле, а в run_generator 
        # передаем уже все полученные значения как параметры - кода, который не относится к логике
        # метода и при этом его можно извлечь, быть в методе не должно

        self.full_file_name: str = self.name_file + '.' + self.format_file
        # Создадим список с фейковыми данными.
        self.data: list = []
        # ↑↑↑ Неполная аннотация типа
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

        # ↑↑↑ Лучше разделить это на методы генерации списка и метод сбора сгенерированных данных
        # в список списков, потому что оба этих метода будут обладать реюзабельностью
        # (e.g. создать список из другого метода с генерацией данных, сделать несколько вариантов генерации)


        if self.format_file == 'xlsx':
            workbook = ex.Workbook(self.full_file_name)
            # ↑↑↑ Нет аннотации типа
            worksheet = workbook.add_worksheet()
            # ↑↑↑ Нет аннотации типа
            for row in range(int(self.number_of_strings)):
                if (row != 0) and (row % 65530 == 0):
                    worksheet = workbook.add_worksheet()
                
            # ↑↑↑ Постоянная проверка на создание нового worksheet серьезно 
            # замедляет цикл без причины, зная итоговое число строк мы можем 
            # вычислить, сколько раз нам нужно создать worksheet для заданного лимита

                col: int = 0
                for data_in_table_cell in self.data[row]:
                    worksheet.write(row - 65530 * (row // 65530), col,
                                    data_in_table_cell)
                    col += 1
                # ↑↑↑ В этой библиотеке есть метод worksheet.write_row(), 
                # который делает работу цикла data_in_table_cell в 1 строку

            workbook.close()
        elif self.format_file == 'csv':
            with open(self.full_file_name, 'w') as f:
                writer = csv.writer(f)
                # ↑↑↑ Нет аннотации типа
                for row in self.data:
                    writer.writerow(row)
        print(f'Файл {self.full_file_name} создан')

        # ↑↑↑ Создание файла с данными нужно вынести 
        # в отдельный класс-сохранялку файлов (в отдельном .py)

    def make_archieve(self):
        # Создает архив из файла.
        # Просим пользователя ввести данные и проверяем их.

        # ↑↑↑ Вместо комментов должен быть докстринг

        # ↑↑↑ Этот метод должен быть вынесен в отдельный
        # класс архивации файлов, в отдельном .py

        print('Теперь создадим архив нашего файла!')
        self.arc_name = input('Введи имя архива или оставь это поле пустым: ')
        # ↑↑↑ Нет аннотации типа
        if self.arc_name == '':
            self.arc_name = f'archive_of_{self.name_file}'
        self.arc_format = input('Введи формат архива: ')
        # ↑↑↑ Нет аннотации типа
        try:
            self.arc_format = self.arc_formats[self.arc_format]
        except KeyError:
            raise KeyError('Такого формата нет в базе')
        self.archieve_name: str = self.arc_name + '.' + self.arc_format
        self.archieve_size = input('Введи предельный размер тома архива: ')
        # ↑↑↑ Нет аннотации типа
        try:
            int(self.archieve_size)
        except ValueError:
            raise ValueError('Необходимо ввести целое число')
        
        
        
        # ↑↑↑ Все то же самое, что говорилось про UI и валидацию выше

        # Cоздаем архив файла
        with zipfile.ZipFile(self.archieve_name, mode='w') as archive:
            archive.write(self.full_file_name)
        # Разделяем архив на тома
        cur_vol: int = 1  # текущий номер тома
        written: int = 0  # сколько байт записали

        # ↑↑↑ Давай переменным такие имена, чтобы 
        # их назначение было понятно без комментов
        
        with open(self.archieve_name, 'rb') as src:
            while True:
                output_name = (f'{self.arc_name}{str(cur_vol)}'
                               f'.{self.arc_format}')
                # ↑↑↑ Нет аннотации типа
                output = open(output_name, 'wb')
                # ↑↑↑ Нет аннотации типа
                while written < (int(self.archieve_size) * self.BYTES_IN_KB):
                    data = src.read(int(self.archieve_size) * self.BYTES_IN_KB)
                    # ↑↑↑ Нет аннотации типа
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
                # ↑↑↑ Нет аннотации типа
                archive.write(cur_name)
                os.remove(cur_name)
        print('Процесс архивирования завершен.')

        # ↑↑↑ Это все надо переписать с использованием буфера (io.BytesIO)
        # плюс не реализована проверка, нужно ли юзеру вообще делить архив

    def add_archieve_format(self, format: str):
        # Добавляет новый формат архива.
        self.arc_formats[format] = format

    # ↑↑↑ Так же идет в файл с классом архивации


def main():
    A = Generator()
    A.add_archieve_format('gzip')
    A.run_generator()
    A.make_archieve()


main()
print('done')
