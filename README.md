# EXCEL-Python-scripts

Скрипты для обработки excell таблицы с большим количесвтом данных
Был написан для большей продуктивности бухгалтерского отдела
Так как это была первая боевая задача и срок стоял один рабочий день, для удобства я разделил скрипт на три части(По возможности будет добавлена оптимищация)
import openpyxl.py - разделяет нужные ячейки по строкам
count openpyxl.py - считает на сколько строк поделится ячейка
add line openpyxl.py - добавляет пустые строки для разделенной ячейки

## Установка

1. Клонируйте репозиторий
2. Перейдите в директорию проекта: `cd EXCEL-Python-scripts`
3. Установите зависимости: `pip install openpyxl`

## Использование
add line openpyxl.py - добавляет пустые строки для разделенной ячейки
```python 
#Открываем файл Excel
workbook = openpyxl.load_workbook('Ваш_файл.xlsx')

#Выбираем нужный лист
sheet = workbook['Имя_вашего_листа']

Создаем новый лист для вставки строк
new_sheet = workbook.create_sheet('Новый список', 1)

#указываем в row[индекс столбца с ячейкой]

for row in sheet.iter_rows(min_row=1, min_col=1, max_col=36):
    # Получаем значение ячейки, которая содержит количество пустых строк для добавления
    cell = row[3]  # Например, пятый столбец (индекс 4) ```

count openpyxl.py - считает на сколько строк поделится ячейка
```python
#Открываем файл Excel
workbook = openpyxl.load_workbook('Ваш_файл.xlsx')

#Выбираем нужный лист
sheet = workbook['Имя_вашего_листа']

#Указываем нужные ячейки
#Проходимся по каждой строке в таблице
for row in sheet.iter_rows(min_row=2, min_col=1, max_col=9):
    # Получаем значение ячейки, в которой нужно подсчитать количество строк данных
    cell_to_count = row[6]  # Например, третий столбец (индекс 2)
    
    # Получаем значение ячейки, в которую нужно вставить количество строк данных
    cell_to_insert = row[1]  # Например, четвертый столбец (индекс 3) ```

import openpyxl.py - разделяет нужные ячейки по строкам
pre> ```python #Открываем файл Excel
workbook = openpyxl.load_workbook('Ваш_файл.xlsx')

#Выбираем нужный лист
sheet = workbook['Имя_вашего_листа']

#Создаем новый лист для обновленных данных
updated_sheet = workbook.create_sheet('Обновленные данные')

#Указываем нужный диапозон для разделения ячеек
#min_row - индекс строки с которой начнется разделение
#min_col и max_col - диапозон столбцов с нужными ячейками

for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
    for cell in row:
        # Получаем значение ячейки
        value = cell.value ``` </pre>