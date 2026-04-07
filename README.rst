Описание
-----------
Менеджер для работы с Excel файлами

Установка пакетом
-----------
Для локальной разработки::
    pip install -e packages/excel_manager
Для обычной установки через requirements.txt::
    excel_manager @ git+https://github.com/dkramorov/excel_manager.git


Импорт
-----------
Проверка::
    from managers.excel_manager import ExcelManager


Удаление
-----------
Удалить пакет::
    pip uninstall excel_manager

Для создания пакета
https://docs.python.org/3.10/distutils/introduction.html#distutils-simple-example
https://docs.python.org/3.10/distutils/sourcedist.html
::
    python setup.py sdist




