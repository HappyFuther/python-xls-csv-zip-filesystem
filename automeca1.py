import xlrd
import zipfile
import csv
from os import listdir
from os.path import isfile, join

print('Searching and extracting zip files')
zipNum = 0
for f in listdir('./'):
  if f.endswith(".zip"):
    zipNum+= 1
    print(str(zipNum)+'-------'+str(f))
    archive = zipfile.ZipFile(str(f), 'r')
    xlsNum = 0
    for afile in archive.namelist():
        if afile.startswith('TecDAX_ICR'):
            xlsNum+= 1
            archive.extract(afile, 'analyze')
            xlsname = str(afile)
            print('\t' + str(xlsNum) + '+++++++' + xlsname)
print('Finished extracting excel files')

print('Start filtering in excel files')
out = csv.writer(open("result.csv", "w", newline=''), delimiter=',', quoting=csv.QUOTE_ALL)
for f in listdir('./analyze'):
  if f.endswith(".xls"):
    try:
        workbook = xlrd.open_workbook('analyze/' + str(f))
        print('Reading file : ' + str(f))

        x = str(f).split('.')
        # Select the worksheet
        worksheet = workbook.sheet_by_name('Data')
        num_rows = 0
        for row in range(worksheet.nrows):
            cell_value = worksheet.cell(row, 0).value
            if cell_value.upper() == "TDXP":
                out.writerow([str(f), x[1], worksheet.cell(row, 3).value, worksheet.cell(row, 4).value, worksheet.cell(row, 5).value, worksheet.cell(row, 25).value])
    except Exception as e:
        print('!File Read Error:' + str(f))
print('Completed : result file is result.csv')