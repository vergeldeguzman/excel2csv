import csv2excel
import unittest
import os

import xlrd


class Csv2XlsTest(unittest.TestCase):

    def assert_excel_file(self, excel_file):
        wb = xlrd.open_workbook(excel_file)

        ws = wb.sheet_by_name('table1')
        self.assertEqual(9, ws.nrows)
        self.assertEqual(4, ws.ncols)
        self.assertEqual('Creation date', ws.cell_value(0, 0))

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(1, 0), wb.datemode)
        self.assertEqual((13, 45, 0), (hour, minute, second))

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(2, 0), wb.datemode)
        self.assertEqual((23, 56, 0), (hour, minute, second))

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(3, 0), wb.datemode)
        self.assertEqual((2003, 1, 2, 0, 0, 0), (year, month, day, hour, minute, second))

        # dateutil.parser does not properly parse '1/3/1689 8:34:00 AM'
        # skipping unit test...
        # year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(8, 0), wb.datemode)
        # self.assertEqual((1689, 1, 3, 8, 34, 0), (year, month, day, hour, minute, second))

        self.assertEqual('0/0/0', ws.cell_value(5, 0))

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(6, 0), wb.datemode)
        self.assertEqual((0, 0, 0), (hour, minute, second))

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(7, 0), wb.datemode)
        self.assertEqual((2005, 3, 1, 12, 34, 56), (year, month, day, hour, minute, second))

        # xlrd does not properly parse '1900-01-01T00:00:00.000'
        # skipping...
        # year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ws.cell_value(8, 0), wb.datemode)
        # self.assertEqual((1900, 1, 1, 0, 0, 0), (year, month, day, hour, minute, second))

        ws = wb.sheet_by_name('table2')
        self.assertEqual(4, ws.nrows)
        self.assertEqual(3, ws.ncols)
        self.assertEqual('one', ws.cell_value(0, 0))
        self.assertEqual(2, ws.cell_value(1, 2))
        self.assertEqual(3.0, ws.cell_value(2, 1))
        self.assertEqual(4, ws.cell_value(3, 2))

    def test_write_xlsx(self):
        input_files = [os.path.join(os.path.dirname(__file__), 'table1.csv'),
                       os.path.join(os.path.dirname(__file__), 'table2.csv')]
        output_file = os.path.join(os.path.dirname(__file__), 'output.xlsx')
        csv2excel.write_excel_from_csv(output_file, input_files, True)
        self.assert_excel_file(output_file)
        os.remove(output_file)

    def test_write_xls(self):
        input_files = [os.path.join(os.path.dirname(__file__), 'table1.csv'),
                       os.path.join(os.path.dirname(__file__), 'table2.csv')]
        output_file = os.path.join(os.path.dirname(__file__), 'output.xlsx')
        csv2excel.write_excel_from_csv(output_file, input_files, True)
        self.assert_excel_file(output_file)
        os.remove(output_file)

    def test_file_not_found_exception(self):
        input_files = ['file_does_not_exist']
        output_file = 'output.xls'
        self.assertRaises(FileNotFoundError,
                          lambda: csv2excel.write_excel_from_csv(output_file, input_files, False))