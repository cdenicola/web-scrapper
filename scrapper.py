import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

#Configurable Data
web_address = "https://pokemondb.net/pokedex/all" # Link from which to pull file
table_class = 'data-table' # Use inspect element on the table, then enter (one of) the class here
dest_filename = 'pokedex' # Output .xlsx destination file


print("Pinging", web_address)
response = requests.get(web_address)

# Check for valid response
if response.status_code > 300 or response.status_code < 200:
    print("Possible Error, response code:", response.status_code)
    exit()
else:
    print("Received Valid Response")

# scrape tables
soup = BeautifulSoup(response.content, 'html.parser')
stat_table = soup.find_all('table', class_=table_class)

# Generate excel sheet
wb = Workbook()
ws1 = wb.active

# Sets workbook title as
try:
    ws1.title = stat_table[0]['id']
except LookupError:
    ws1.title = table_class

# Scape tables, save data in excel sheet
for i, row in enumerate(stat_table[0].find_all('tr')):
    for j, cell in enumerate(row.find_all('td')):  # place into excel here
        ws1.cell(row=i + 1, column=j + 1, value=cell.text)

# Save file
wb.save(filename=dest_filename + ".xlsx")
print("Successful generated ", dest_filename, ".xls", sep="")
