# excel2csv

Script to convert Excel xls, xlsx or xml file to csv. Optionally, user can specify xml namespace 
which defaults to `urn:schemas-microsoft-com:office:spreadsheet`. 

## Running the tests

```
python3 -m unittest tests/test_xls2csv.py
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

## Requirements

    python 3.5
    lxml
    xlrd

## Example run

Convert xlsx file to csv

```
python3 excel2csv.py -i workbook.xlsx
```

Convert Excel xml file to csv

```
python3 excel2csv.py -i workbook.xml
```
