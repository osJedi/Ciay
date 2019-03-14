import gspread
import time
import os
import sys
import re
from oauth2client.service_account import ServiceAccountCredentials

if (len(sys.argv) > 1 and "kingfisher H3 M3 M3N".find(sys.argv)):
    subsheet = sys.argv[1]
else:
    #print("ERROR IN PARAMETER"); exit 1
    subsheet = 'kingfisher'


if (len(sys.argv) > 2):
    path_to_csv = sys.argv[2]
else:
    path_to_csv = "data_result_OK.csv"


scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('CIProject-752eb8bdfba6.json', scope)

gc = gspread.authorize(credentials)

sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1oUW4Gyd5hcmVW5q8D3l3jAx-NBCly778SlLECfsErew')
#worksheet = sheet.get_worksheet(1)
worksheet = sheet.worksheet('kingfisher')

# GET number of column for NEW RECORD
control_string_free_column = worksheet.get_all_values()[2:3][0]
next_free_column = len(",".join(filter(lambda i: not i=='', control_string_free_column)).split(',')) + 1
print(next_free_column)

# Get the RELEASE VERSION
control_string_release = worksheet.get_all_values()[0:1][0]
release_column = len(",".join(filter(lambda i: not i=='', control_string_release)).split(',')) * 4 - 11
release = worksheet.cell(1, release_column).value
print(release)


def letter_calc(number):
    alphabet = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I", 10: "J", 11: "K", 12: "L", 13: "M", 14: "N", 15: "O", 16: "P", 17: "Q", 18: "R", 19: "S", 20: "T", 21: "U", 22: "V", 23: "W", 24: "X", 25: "Y", 26: "Z" }
    fist = number % 26
    if (number > 26):
        fist = number // 26
        second = number % 26
        return (str(alphabet[fist]) + str(alphabet[second]))
    else:
        return (str(alphabet[fist]))


# Get CONTROL_LIST from document
control_list = []
st = letter_calc(next_free_column)
fn = letter_calc(next_free_column + 3)
sss = st + "3" + ":" + fn + "34"
control_list = worksheet.range(sss)
print(type(con_list))

f = open(path_to_csv, 'r')
i = 4
for cell in control_list:
    if (i > 3):
        for line in f:
            test_name = line.split(',')[0]
            test_frames = line.split(',')[1]
            test_time = line.split(',')[2]

            cell_list = worksheet.findall(test_name)
            coordinates = str(cell_list[0])[7:]
            test_row = re.findall(r'\d+', coordinates)[0]

            # Create FORMULA FOR FRAMES
            formula1 = "=" + letter_calc(next_free_column) + str(test_row) + "/" + letter_calc(next_free_column - 4) + str(test_row) + "-1"

            # Create FORMULA FOR TIME
            formula2 = "=" + letter_calc(next_free_column + 2) + str(test_row) + "/" + letter_calc(next_free_column - 4) + str(test_row) + "-1"

            break

    print(cell)
    if (i == 0):
        cell.value = test_frames
    if (i == 1):
        cell.value = formula1
    if (i == 2):
        cell.value = test_time
    if (i == 3):
        cell.value = formula2
    i += 1

print(control_list)
worksheet.update_cells(control_list, value_input_option='USER_ENTERED')





f.close()



# print(worksheet)
# print(worksheet.get_all_values())
# print(worksheet.range('AR3:AS3'))
# val = worksheet.cell(3, 44).value
# print(val)


# for i in range(1, 100):
#     val = worksheet.cell(i, 3).value
#     #if (val == "" ):
#     print(i)
#     print(val)
#     #break
