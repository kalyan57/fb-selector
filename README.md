# fb-selector

For English description look down.

## [ОПИСАНИЕ]

Скрипт делает выборку табельного номера, номера и серии действующего пропуска для каждого сотрудника указанной организации или подразделения.

## [ТРЕБОВАНИЯ]

Скрипт написан на Python 3.7.0
Перед использованием требуется установка дополнительных библиотек командой
```
pip install fdb pandas xlwt argparse pywin32
```

## [ИСПОЛЬЗОВАНИЕ]

Скрипт принимает название организации ИЛИ подразделения в качестве входных параметров и создает выходной Excel файл с найденными данными.
Имя выходного файла также может быть указано в качестве входного параметра.
Список всех параметров, принимаемых скриптом. может быть получен командой:
```
> python .\fb-selector.py -h
usage: fb-selector.py [-h] [-d DEP | -o ORG] [-out OUT]

fb-selector allows us to select PASS CARD id and numbers of employees of a
specified department or organization

optional arguments:
  -h, --help            show this help message and exit
  -d DEP, --depatrment DEP
                        Submit department name
  -o ORG, --organization ORG
                        Submit organization name
  -out OUT, --output OUT
                        Submit output Excel file name. (default is out.xls)
```

Если имя выходного файла не указано при запуске, то используется название по умолчанию - out.xls, файл будет создан в директории, из которой был запущен скрипт. Указать имя выходного файла можно командой
```
> python .\fb-selector.py -d "department name" -out 'c:\mydir\my_out_file.xls'
```
Для выборки пропусков сотрудников по названию подразделения используйте команду
```
> python .\fb-selector.py -d "department name"
```
Для выборки пропусков сотрудников по названию организации используйте команду
```
> python .\fb-selector.py -o "organization name"
```
Поиск регистрозависим, но можно искать по части названия. Однако для увеличения точности и во избежание попадания в выборку данных из структурных подразделений с похожим названием лучше название указывать точно и целиком.

