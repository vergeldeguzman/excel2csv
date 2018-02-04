import excel2csv
import unittest
import os
import filecmp


class Excel2CsvTest(unittest.TestCase):

    def test_parse_xml_1table(self):
        input_file = os.path.join(os.path.dirname(__file__), '1table.xml')
        excel2csv.parse_xml(input_file, excel2csv.DEFAULT_XML_NAMESPACE)

        output_file = '1table_Sheet1.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '1table_Sheet1.expected.xml.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

    def test_parse_xls_1table(self):
        input_file = os.path.join(os.path.dirname(__file__), '1table.xlsx')
        excel2csv.parse_xls(input_file)

        output_file = '1table_Sheet1.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '1table_Sheet1.expected.xlsx.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

    def test_parse_xml_2sheets(self):
        input_file = os.path.join(os.path.dirname(__file__), '2sheets.xml')
        excel2csv.parse_xml(input_file, excel2csv.DEFAULT_XML_NAMESPACE)

        output_file = '2sheets_Sheet1.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '2sheets_Sheet1.expected.xml.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

        output_file = '2sheets_Sheet2.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '2sheets_Sheet2.expected.xml.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

    def test_parse_xls_2sheets(self):
        input_file = os.path.join(os.path.dirname(__file__), '2sheets.xlsx')
        excel2csv.parse_xls(input_file)

        output_file = '2sheets_Sheet1.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '2sheets_Sheet1.expected.xlsx.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

        output_file = '2sheets_Sheet2.csv'
        expected_file = os.path.join(os.path.dirname(__file__), '2sheets_Sheet2.expected.xlsx.csv')
        self.assertTrue(filecmp.cmp(output_file, expected_file))
        os.remove(output_file)

    def test_parse_xls_ambiguous_date_exception(self):
        input_file = os.path.join(os.path.dirname(__file__), 'ambiguousdate.xlsx')
        self.assertRaises(excel2csv.Excel2CsvException,
                          lambda: excel2csv.parse_xls(input_file))

    def test_parse_xml_exception(self):
        input_file = os.path.join(os.path.dirname(__file__), '1table.xlsx')
        self.assertRaises(excel2csv.Excel2CsvException,
                          lambda: excel2csv.parse_xml(input_file, excel2csv.DEFAULT_XML_NAMESPACE))

    def test_parse_xml_file_does_not_exist(self):
        input_file = 'file_does_not_exist'
        self.assertRaises(OSError,
                          lambda: excel2csv.parse_xml(input_file, excel2csv.DEFAULT_XML_NAMESPACE))

    def test_parse_xls_exception(self):
        input_file = os.path.join(os.path.dirname(__file__), '1table.xml')
        self.assertRaises(excel2csv.Excel2CsvException,
                          lambda: excel2csv.parse_xls(input_file))

    def test_parse_xls_file_does_not_exist(self):
        input_file = 'file_does_not_exist'
        self.assertRaises(OSError,
                          lambda: excel2csv.parse_xls(input_file))


if __name__ == '__main__':
    unittest.main()
