#!/usr/bin/env python3

import os
import sys
import argparse
import csv
import datetime
from lxml import etree

import xlrd

DEFAULT_XML_NAMESPACE = 'urn:schemas-microsoft-com:office:spreadsheet'


class Xls2CsvException(Exception):
    pass


def parse_xml(xml_file, ns):
    try:
        tree = etree.parse(xml_file)
        root = tree.getroot()
        worksheets_xml = root.findall(etree.QName(ns, 'Worksheet').text)
        for worksheet_xml in worksheets_xml:
            rows = []

            table_xml = worksheet_xml.find(etree.QName(ns, 'Table').text)
            rows_xml = table_xml.findall(etree.QName(ns, 'Row').text)
            for row_xml in rows_xml:
                row = []
                cells_xml = row_xml.findall(etree.QName(ns, 'Cell').text)
                for cell_xml in cells_xml:
                    data_xml = cell_xml.find(etree.QName(ns, 'Data').text)
                    row.append(data_xml.text)
                rows.append(row)

            base_name = os.path.splitext(os.path.basename(xml_file))[0]
            worksheet_name = worksheet_xml.attrib.get(etree.QName(ns, 'Name').text)
            csv_file = base_name + '_' + worksheet_name + '.csv'
            write_to_csv(rows, csv_file)
    except etree.XMLSyntaxError:
        raise Xls2CsvException('Cannot parse xml file: ' + xml_file)


def get_cell_value(workbook, worksheet, row, col):
    cell_type = worksheet.cell_type(row, col)
    cell_value = worksheet.cell_value(row, col)
    if cell_type == xlrd.XL_CELL_DATE:
        dt_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
        cell_dt = datetime.datetime(dt_tuple[0], dt_tuple[1], dt_tuple[2], dt_tuple[3], dt_tuple[4], dt_tuple[5])
        cell_value = cell_dt.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3]

    return str(cell_value)


def parse_xls(xls_file):
    try:
        book_xls = xlrd.open_workbook(xls_file)
        for sheet_xls in book_xls.sheets():

            rows = []
            for row_idx in range(sheet_xls.nrows):
                row = []
                for col_idx in range(sheet_xls.ncols):
                    cell_value = get_cell_value(book_xls, sheet_xls, row_idx, col_idx)
                    row.append(cell_value)
                rows.append(row)

            base_name = os.path.splitext(os.path.basename(xls_file))[0]
            csv_file = base_name + '_' + sheet_xls.name + '.csv'
            write_to_csv(rows, csv_file)
    except xlrd.biffh.XLRDError:
        raise Xls2CsvException('Cannot parse excel file: ' + xls_file)


def write_to_csv(rows, csv_file):
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)


def parse_arg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input-file',
                        required=True,
                        help='input excel file')
    parser.add_argument('-x', '--xml',
                        action='store_true',
                        help='parse input file as xml')
    parser.add_argument('-n', '--xml-namespace',
                        default=DEFAULT_XML_NAMESPACE,
                        help='namespace for excel xml file')

    args = parser.parse_args()
    return args


def main(argv):
    try:
        args = parse_arg()
        if args.xml:
            parse_xml(args.input_file, args.xml_namespace)
        else:
            parse_xls(args.input_file)
    except Exception as e:
        print(e, file=sys.stderr)
        exit(1)


if __name__ == "__main__":
    main(sys.argv)
