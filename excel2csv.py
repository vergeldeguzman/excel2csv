#!/usr/bin/env python3

import os
import sys
import argparse
import csv
from datetime import datetime, time
from lxml import etree

import xlrd

DEFAULT_XML_NAMESPACE = 'urn:schemas-microsoft-com:office:spreadsheet'


class Excel2CsvException(Exception):
    pass


def get_cell_value_from_xml(xml_cell, ns):
    xml_cell_data = xml_cell.find(etree.QName(ns, 'Data').text)
    cell_data = xml_cell_data.text
    cell_type = xml_cell_data.attrib.get(etree.QName(ns, 'Type').text)
    if cell_type == 'DateTime':
        dt = datetime.strptime(cell_data, '%Y-%m-%dT%H:%M:%S.%f')
        if dt.year == 1899 and dt.month == 12 and dt.day == 31: # no date
            cell_data = dt.strftime('%H:%M:%S.%f')[:-3]
    return cell_data


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
                    cell_value = get_cell_value_from_xml(cell_xml, ns)
                    row.append(cell_value)
                rows.append(row)

            base_name = os.path.splitext(os.path.basename(xml_file))[0]
            worksheet_name = worksheet_xml.attrib.get(etree.QName(ns, 'Name').text)
            csv_file = base_name + '_' + worksheet_name + '.csv'
            write_to_csv(rows, csv_file)
    except etree.XMLSyntaxError:
        raise Excel2CsvException('Cannot parse xml file: ' + xml_file)


def get_cell_value_from_excel(workbook, worksheet, row, col):
    cell_type = worksheet.cell_type(row, col)
    cell_value = worksheet.cell_value(row, col)
    if cell_type == xlrd.XL_CELL_DATE:
        try:
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
            if year == 0 and month == 0 and day == 0: # no date specified
                cell_dt = time(hour, minute, second)
                cell_value = cell_dt.strftime('%H:%M:%S.%f')[:-3]
            else:
                cell_dt = datetime(year, month, day, hour, minute, second)
                cell_value = cell_dt.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3]
        except xlrd.xldate.XLDateAmbiguous:
            # https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year
            raise Excel2CsvException('Ambiguous date on worksheet {} row {} column {}'.format(worksheet.name, row, col))

    return str(cell_value)


def parse_xls(excel_file):
    try:
        book_xls = xlrd.open_workbook(excel_file)
        for sheet_xls in book_xls.sheets():

            rows = []
            for row_idx in range(sheet_xls.nrows):
                row = []
                for col_idx in range(sheet_xls.ncols):
                    cell_value = get_cell_value_from_excel(book_xls, sheet_xls, row_idx, col_idx)
                    row.append(cell_value)
                rows.append(row)

            base_name = os.path.splitext(os.path.basename(excel_file))[0]
            csv_file = base_name + '_' + sheet_xls.name + '.csv'
            write_to_csv(rows, csv_file)
    except xlrd.biffh.XLRDError:
        raise Excel2CsvException('Cannot parse excel file: ' + excel_file)


def write_to_csv(rows, csv_file):
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)


def parse_arg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input-file',
                        required=True,
                        help='input excel file, supported: xls, xlsx or xml')
    parser.add_argument('-n', '--xml-namespace',
                        default=DEFAULT_XML_NAMESPACE,
                        help='namespace for excel xml file')

    args = parser.parse_args()
    return args


def main(argv):
    try:
        args = parse_arg()
        dummy, file_extension = os.path.splitext(os.path.basename(args.input_file))
        if file_extension.lower() == '.xml':
            parse_xml(args.input_file, args.xml_namespace)
        else:
            parse_xls(args.input_file)
    except Exception as e:
        print(e, file=sys.stderr)
        exit(1)


if __name__ == "__main__":
    main(sys.argv)
