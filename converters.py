import sys, codecs, datetime

class CHConverter(object):
    def __init__(self, input_file, output_file, source_encoding="WINDOWS-1251", raw_date_fields = None, empty_numbers=None, empty_strings=None):
        
        self.raw_date_fields = raw_date_fields # Массив, начиная с 0, колонок, для которых даты вида 42250 надо привести в нормальные
        self.empty_numbers = empty_numbers # Если это число может быть пропущено, то ставим 0
        self.empty_strings = empty_strings # Если пустая строка, то ставим " "
        self.input_file = input_file
        self.output_file = output_file
        self.source_encoding = source_encoding


    def convert(self):
        try:

            if not self.raw_date_fields is None:
                start_date = datetime.datetime.strptime("01/01/1990", "%d/%m/%Y")
                
            output_file = codecs.open(self.output_file, "wb", "utf8")
            print("Загружаю рабочую книгу")
            if self.input_file[:-3] == 'LSX':
                from openpyxl import load_workbook
                wb = load_workbook(self.input_file)
                ws = wb.active
            else:
                import xlrd
                wb = xlrd.open_workbook(self.input_file)
                sheet = wb.sheet_by_index(0)
                ws = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
                
            print("Экспортирую данные")  
            for row_num, row in enumerate(ws):
                str_val = ""
                for pos, cell in enumerate(row):

                    try:
                        cell_value = cell.value
                    except AttributeError as e:
                        cell_value = cell
                        
                    
                    if self.raw_date_fields and pos in self.raw_date_fields:
                        try:
                            val = (start_date + datetime.timedelta(days=int(cell_value)-3)).strftime("%Y-%m-%d")
                        except ValueError: # Если это заголовок
                            val = cell_value
                        except OverflowError as e:
                            print("Ошибка: '{0}' Дата: '{1}'".format(e, cell_value))
                            sys.exit()
                    else:
                        val = cell_value
                    if not val:
                        if self.empty_strings and pos in self.empty_strings:
                            val = ""
                        elif self.empty_numbers and pos in self.empty_numbers:
                            val = 0
                        
                    str_val += str(val) + '\t'
                
                output_file.write(str_val[:-1] + "\n")
            output_file.close()
                
        except FileNotFoundError as e:
            print("Не удалось открыть файл: {0}".format(e))
            
"""
conv = CHConverter("./top/16.06.2016.xls", "./top_res/16_06_2016.txt", empty_numbers = [
        1,6,7,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,
        31,32,33,34,35,36,37,38,39,40,
        41,42,43,44,45,46,47,48,49,50,
        51,52,53,54,55,56,57,58,59,60,
        61,62,63,64,65,66,67,68,69,70,
        71,72,73,74,75,76,77,78,79,80,
        81,82,83,84,85,86,87,88,89,90,
        91,92,93,94,95,96,97,98,99,100,
], empty_strings=[3,4])
"""

conv = CHConverter("./test_src/01_01_2016.xlsx", "./test_res/16_06_2016.txt", empty_numbers = [
        1,6,7,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,
        31,32,33,34,35,36,37,38,39,40,
        41,42,43,44,45,46,47,48,49,50,
        51,52,53,54,55,56,57,58,59,60,
        61,62,63,64,65,66,67,68,69,70,
        71,72,73,74,75,76,77,78,79,80,
        81,82,83,84,85,86,87,88,89,90,
        91,92,93,94,95,96,97,98,99,100,
], empty_strings=[3,4])
    

conv.convert()
