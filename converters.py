import sys, codecs, datetime

class XLSXConverter(object):
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
            from openpyxl import load_workbook
            wb = load_workbook(self.input_file)
            ws = wb.active
            print("Экспортирую данные")  
            for row_num, row in enumerate(ws):
                str_val = ""
                for pos, cell in enumerate(row):
                    if self.raw_date_fields and pos in self.raw_date_fields:
                        try:
                            val = (start_date + datetime.timedelta(days=int(cell.value)-3)).strftime("%Y-%m-%d")
                        except ValueError: # Если это заголовок
                            val = cell.value
                        except OverflowError as e:
                            print("Ошибка: '{0}' Дата: '{1}'".format(e, cell.value))
                            sys.exit()
                    else:
                        val = cell.value
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

        
#conv = XLSXConverter("./31_05_2016.xlsx", "./31_05_2016.txt", raw_date_fields = [13], empty_numbers = [12,17,19], empty_strings=[20])
#conv.convert()
