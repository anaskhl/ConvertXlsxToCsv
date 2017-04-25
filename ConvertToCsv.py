import xlrd
import csv

def convert(s):
    if isinstance(s, str):
        return s.encode("utf-8")
    elif isinstance(s, unicode):
        return s.encode("utf-8")
    elif (s.is_integer()):
		return int(s)
    else:
		return s

quote = {'1': csv.QUOTE_ALL,
         '2': csv.QUOTE_MINIMAL,
		 '3': csv.QUOTE_NONNUMERIC,
		 '4': csv.QUOTE_NONE,
    } 
		
file_location = raw_input("Please enter the file location, preferably .xlsx file: ")	
print('It might take some time depending on how big the excel file is, please wait')	
wb = xlrd.open_workbook(file_location)
sheet_name = raw_input("Please enter the sheet name: ")
sh = wb.sheet_by_name(sheet_name)
csv_file_name = raw_input("Please enter the csv file name with the file extension .csv: ")
print('Please select the quoting option, it is how the csv should quote the valyes')
print('1 - QUOTE_ALL: means that all the values will be quoted')
print('2 - QUOTE_MINIMAL: to only quote those fields which contain special characters such as delimiter, quotechar or any of the characters in lineterminator')
print('3- QUOTE_NONNUMERIC:  to quote all non-numeric fields')
print('4- QUOTE_NONE: to never quote fields. When the current delimiter occurs in output data it is preceded by the current escapechar character. If escapechar is not set, the writer will raise Error if any characters that require escaping are encountered.')
print('Please write only the number corrseponds to the option')

your_csv_file = open(csv_file_name, 'wb')

quote_option = raw_input("Please write number for the quoting option: ")
wr = csv.writer(your_csv_file, quoting = quote[quote_option])
for rownum in xrange(sh.nrows):
	wr.writerow([convert(s) for s in sh.row_values(rownum)])
your_csv_file.close()
print 'done'