# xls2csv

Script to convert Excel xls/xlsx file to csv. In addition, it can convert
Excel workbook saved in xml format. Optionally, use can specify xml namespace 
which defaults to urn:schemas-microsoft-com:office:spreadsheet. 

## Running the tests

```
python3 -m unittest tests/test_xls2csv.py
```

## Usage

```
usage: xls2csv.py [-h] -i INPUT_FILE [-x] [-n XML_NAMESPACE]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FILE, --input-file INPUT_FILE
                        input excel file
  -x, --xml             parse input file as xml
  -n XML_NAMESPACE, --xml-namespace XML_NAMESPACE
                        namespace for excel xml file
```

## Requirements

	python 3.5
    xlrd

## Example run

Convert xlsx file to csv

```
python3 xls2csv.py -i workbook.xlsx
```

Convert Excel xml file to csv

```
python3 xls2csv.py -i another_workbook.xml -x xml
```
