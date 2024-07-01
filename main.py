from glob import glob
import openpyxl

# find all excels
files = glob('./files/*')
counter = 0

for file in files:

    current_excel_file = openpyxl.load_workbook(file)

    sheets = list()

    # add every sheets name to sheets
    for sheet in current_excel_file:
        sheets.append(sheet.title)

    # loop on every exist sheet
    for single_sheet in sheets:
        current_excel_sheet = current_excel_file[single_sheet]
        null_counter = 0

        for row in current_excel_sheet.iter_rows(0,current_excel_sheet.max_row,values_only=True):

            is_ended = False
            for single_value in row:
                print(str(counter)  + ' ' + str(single_value))
                counter += 1

                if single_value == None:
                    null_counter += 1
                else:
                    null_counter -= 1

                if single_value == None and null_counter >= 200:
                    is_ended = True
                    break

            if is_ended == True:
                print('='*100)
                print(counter)
                print(file)
                break
