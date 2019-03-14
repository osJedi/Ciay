import gspread

import os
import sys
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



f = open(path_to_csv, 'r')

for line in f:
    print(line.split(',')[0])






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
