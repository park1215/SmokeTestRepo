__author__= 'Sean Park_ViaSat'

import datetime
import time
import string
import random
import pandas as pd
import xlsxwriter

now = datetime.datetime.now()

currentYear = str(now.year)
currentMonth = str(now.month)
currentDay = str(now.day)

print(currentYear)

print(currentMonth)

print(currentDay)

months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

currentMonth = months[int(currentMonth)-1]

print(currentMonth)

hexdigits = list(string.hexdigits)
del hexdigits[10:16]

print(hexdigits)

randomMac = "AA:BB:CC:"

for x in range(0, 6):
    randomNumber = random.choice(hexdigits)
    randomMac = randomMac + randomNumber
    if x % 2 != 0 and len(randomMac) < 17:
        randomMac = randomMac+":"

# Create a Pandas dataframe from the data
df = pd.DataFrame({'Data':[10, 20, 30, 40, 20, 15, 41, 53, 90]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file
writer.save()


