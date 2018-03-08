# excel2csv / csv2excel

excel2csv is a script to convert Excel xls , xlsx or xml file to csv file. The csv files are generated from each Excel sheet.
For example, Workbook.xlsx has 3 sheets (i.e. Sheet1, Sheet2, Sheet3), the script generates Workbook_Sheet1.csv,
Workbook_Sheet2.csv, Workbook_Sheet3.csv.
 
Optionally, user can specify xml namespace 
which defaults to `urn:schemas-microsoft-com:office:spreadsheet` to generates xml files.

csv2excel is another script to convert csv file to Excel xls file or Excel xlsx file. The script
accept several csv files per Excel sheet. 
 
## Running the tests

```
python3 -m unittest tests\test_excel2csv.py tests\test_csv2excel.py
```

## Usage

```
usage: excel2csv.py [-h] -i INPUT_FILE [-n XML_NAMESPACE]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FILE, --input-file INPUT_FILE
                        input excel file
  -n XML_NAMESPACE, --xml-namespace XML_NAMESPACE
                        namespace for excel xml file
```
```
usage: csv2excel.py [-h] -i INPUT_FILES [INPUT_FILES ...] -o OUTPUT_FILE [-t]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FILES [INPUT_FILES ...], --input-files INPUT_FILES [INPUT_FILES ...]
                        input excel files
  -o OUTPUT_FILE, --output-file OUTPUT_FILE
                        output excel file
  -t, --translate-date  translate date string to excel date
  ```
  
## Requirements

    python 3.5
    lxml
    xlrd
    openpyxl
    dateutils

## Example run

Convert xlsx file to csv file

```
python3 excel2csv.py -i workbook.xlsx
```

Convert Excel xml file to csv file

```
python3 excel2csv.py -i workbook.xml 
```

Convert csv files to xlsx file with date string in Excel date/time format

```
python3 csv2excel.py -i table1.csv -i table2.csv -t -o workbook.xlsx 
```

Convert csv files to xlsx file (date string is kept as string in Excel file)

```
python3 csv2excel.py -i table1.csv -i table2.csv -o workbook.xlsx 
```
