#!/bin/python3

import os
import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pickle
import os.path
from googleapiclient.discovery import build
from google.auth.transport.requests import Request


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('CIProject-752eb8bdfba6.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Renesas_4.1.2.1_Android_8.1_Benchmark_Results").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_records()
print(list_of_hashes)