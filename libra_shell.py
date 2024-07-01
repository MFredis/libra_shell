import openpyxl
import xlrd
from os import listdir

# xlsfiles = [f for f in listdir('.') if f.split('.')[1] == 'xls']
# if len(xlsfiles) == 0:
#     print("No .xls files")
#     exit()

ORDER_FILE = 'Заявка.xls'
FORM_FILE = 'Форма1.xlsx'

print("Reading: " + ORDER_FILE)

# Open the source Excel file
source_file_path = ORDER_FILE
source_workbook = xlrd.open_workbook(source_file_path, encoding_override="cp1251")
source_worksheet = source_workbook.sheet_by_index(0)

# Open the destination Excel file
destination_file_path = FORM_FILE
destination_workbook = openpyxl.load_workbook(destination_file_path)
destination_worksheet = destination_workbook.active

# Define the column to iterate through
columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K']

# Starting and ending row index
start_row = 3
end_row = source_worksheet.nrows

for column in columns:
    # Iterate through the source columns
    for row in range(start_row, end_row + 1):
        # cell_value = source_worksheet[column + str(row)].value
        cell_value = source_worksheet.cell_value(row - 1, ord(column) - ord('A'))
        destination_worksheet[column + str(row)] = cell_value

# Save changes to the destination Excel file
destination_workbook.save('Результат.xlsx')

# Close both workbooks
source_workbook.release_resources()
destination_workbook.close()

print("Data has been successfully written to the destination Excel file.")
