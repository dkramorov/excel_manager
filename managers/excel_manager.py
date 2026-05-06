from io import BytesIO
from openpyxl import load_workbook


class ExcelManager:
    """Менеджер по работе с xlsx файлами
    """
    def __init__(self, file_path: str):
        """Инициализация
           :param file_path: путь к файлу
        """
        self.file_path = file_path
        self.wb = None

    def open_wb(self, data_only: bool = True):
        """Загружаем эксельку
           :param data_only: без формул - только значения
        """
        with open(self.file_path, 'rb') as excel_file:
            self.wb = load_workbook(BytesIO(excel_file.read()), data_only=data_only)
        # Посмотреть листы эксельки: logger.info(wb.sheetnames)
        return self.wb

    def search_header(self, rows, symbols: str = '№'):
        """Поиск строки заголовков
           обход по генератору,
           после вызова мы будем находиться на строке заголовка
           :param rows: генератор для строк
           :param symbols: символы с которых начинается заголовок таблицы
        """
        start_row_index = 0
        for row in rows:
            start_cell_index = 0
            for cell in row:
                value = cell.value
                if value and symbols in str(value):
                    return start_row_index, start_cell_index, row
                start_cell_index += 1
            start_row_index += 1
        return None, None, None

    def accumulate_data(self, row, i: int, data: list):
        """Добавление ячейки в массив данных
           USAGE: data_arr = accumulate_data(row, i+1, result)
           где row - строка, где находимся генератором,
           i + 1 - номер ячейки по которой аккумулируем данные
           result - накопленные данные
           :param row: строка эксельки
           :param i: номер ячейки в строке эксельки
           :param data: массив данных
        """
        value = row[i].value
        if value:
            value = '%s' % value
            value = value.strip()
            if value and not value in data:
                data.append(value)
        return value

    def load_wb(self, required_names: list = None, search_header_symbols: str = None):
        """Получить данные из эксель файла
           :param required_names: список обязательных полей в шапке
           :param search_header_symbols: символы с которых начинается заголовок таблицы
        """
        result = {
            'errors': [],
            'data': [],
        }
        if not required_names:
            required_names = []
        wb = self.open_wb()
        sheet = wb.active

        rows = list(sheet.rows)
        if not len(rows) > 0:
            result['errors'].append('Файл пустой')
            return result

        # Получаем шапку
        if search_header_symbols:
            start_row_index, start_cell_index, row = self.search_header(
                rows=rows, symbols=search_headers_symbols,
            )
            names = [cell.value for cell in row]
        else:
            start_row_index = 0
            start_cell_index = 0
            names = [cell.value for cell in rows[0]]

        for name in required_names:
            if not name in names:
                result['errors'].append('обязательное поле %s' % name)

        if result['errors']:
            return result

        for i, row in enumerate(rows):
            if i < start_row_index:
                continue
            obj = {}
            for j, cell in enumerate(row):
                if j < start_cell_index:
                    continue
                key = names[j - start_cell_index]
                if not key:
                    continue
                value = str(cell.value).strip() if cell.value != None else ''
                obj[key] = value
            result['data'].append(obj)
        return result
