На данный момент умеет файлы формата XLSX приводить в "нужный" формат для заливки в clickhouse - меняет переводы строк, кодировку файла, заменяет пустые ячейки на 0 или " ", своеобразные даты Excel - 43320 в YYYY-MM-DD
Разделителем ставим табуляции, так как для этого не надо экранировать строки
Заголовки оставляем, в БД они записаны не будут

Пример использования - 

```
conv = XLSXConverter(INPUT_FILE, DEST_FILE, raw_date_fields = [13], empty_numbers = [12,17,19], empty_strings=[20])
conv.convert()
```

где 

**raw_date_strings** - номер(а) колонки (начиная с нуля), где стоит дата в формате 43123 

**empty_numbers** - номер(а) колонок, где указываются цифры, но может быть и пустая колонка, тогда там поставим 0

**empty_strings** - номер(а) колонок, где пустая ячейка будет заменена на " "


Заливка в базу может быть осуществлена командой
```
cat ~/YOUR_FILE_NAME | clickhouse-client --query="INSERT INTO YOUR_TABLE_NAME FORMAT TabSeparatedWithNames"
```

Для того, что бы променять пачку файлов, можно воспользоваться внешним скриптом
(исходные файлы лежат в ./transf, новые будут складываться в ./res)

import re
import subprocess
from os import listdir
from os.path import isfile, join
from converters import XLSXConverter

patt = ".+(\d{2}_\d{2}_\d{4}).+"
mypath = "./transf"
for f in [f for f in listdir(mypath) if isfile(join(mypath, f))]:
    source_file = "./transf/"+f
    dest_file = "./res/" + re.findall(patt, f)[0]
    print("Обработка файла {0}".format( f) )
    conv = XLSXConverter(source_file, dest_file, raw_date_fields = [13])
    conv.convert()


