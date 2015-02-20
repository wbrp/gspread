# -*- coding: utf-8 -*-
import os
import re
import time
import random
import hashlib
import unittest
import ConfigParser
import itertools

import gspread


def _create_worksheet(spreadsheet):
    cur_sheet_name = 'Test In Progress - %s' % time.asctime()
    spreadsheet.add_worksheet(cur_sheet_name, 16, 16)
    return spreadsheet.worksheet(cur_sheet_name)


class GspreadTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        creds_filename = "tests.config"
        try:
            config_filename = os.path.join(
                os.path.dirname(__file__), creds_filename)
            config = ConfigParser.ConfigParser()
            config.readfp(open(config_filename))
            email = config.get('Google Account', 'email')
            password = config.get('Google Account', 'password')
            cls.config = config
            cls.gc = gspread.login(email, password)
        except IOError:
            msg = "Can't find %s for reading google account credentials. " \
                  "You can create it from %s.example in tests/ directory."
            raise Exception(msg % (creds_filename, creds_filename))


class ClientTest(GspreadTest):
    """Test for gspread.client."""

    def test_open(self):
        title = self.config.get('Spreadsheet', 'title')
        spreadsheet = self.gc.open(title)
        self.assertTrue(isinstance(spreadsheet, gspread.Spreadsheet))

    def test_no_found_exeption(self):
        noexistent_title = "Please don't use this phrase as a name of a sheet."
        self.assertRaises(gspread.SpreadsheetNotFound,
                          self.gc.open,
                          noexistent_title)

    def test_open_by_key(self):
        key = self.config.get('Spreadsheet', 'key')
        spreadsheet = self.gc.open_by_key(key)
        self.assertTrue(isinstance(spreadsheet, gspread.Spreadsheet))

    def test_open_by_url(self):
        url = self.config.get('Spreadsheet', 'url')
        spreadsheet = self.gc.open_by_url(url)
        self.assertTrue(isinstance(spreadsheet, gspread.Spreadsheet))

    def test_openall(self):
        spreadsheet_list = self.gc.openall()
        for s in spreadsheet_list:
            self.assertTrue(isinstance(s, gspread.Spreadsheet))


class WorksheetDeleteTest(GspreadTest):
    """Test whether deleting worksheets works."""

    @classmethod
    def setUpClass(cls):
        super(WorksheetDeleteTest, cls).setUpClass()
        title = cls.config.get('Spreadsheet', 'title')
        cls.spreadsheet = cls.gc.open(title)
        cls.ws1 = cls.spreadsheet.add_worksheet('delete_ws_1', 1, 1)
        cls.ws2 = cls.spreadsheet.add_worksheet('delete_ws_2', 100, 100)

    def test_delete_multiple_worksheets(self):
        self.spreadsheet.del_worksheet(self.ws1)
        self.spreadsheet.del_worksheet(self.ws2)


class SpreadsheetTest(GspreadTest):
    """Various tests for gspread.Spreadsheet."""

    @classmethod
    def setUpClass(cls):
        super(SpreadsheetTest, cls).setUpClass()
        title = cls.config.get('Spreadsheet', 'title')
        cls.spreadsheet = cls.gc.open(title)

    def test_properties(self):
        self.assertEqual(self.config.get('Spreadsheet', 'key'),
                         self.spreadsheet.id)
        self.assertEqual(self.config.get('Spreadsheet', 'title'),
                         self.spreadsheet.title)

    def test_sheet1(self):
        sheet1 = self.spreadsheet.sheet1
        self.assertTrue(isinstance(sheet1, gspread.Worksheet))

    def test_get_worksheet(self):
        sheet1 = self.spreadsheet.get_worksheet(0)
        self.assertTrue(isinstance(sheet1, gspread.Worksheet))


class WorksheetPropertiesTest(GspreadTest):
    """Test if known properties are equal to downloaded properties."""

    def test_properties(self):
        conf = self.config
        worksheet_title = conf.get('Worksheet Properties', 'title')
        spreadsheet_title = conf.get('Spreadsheet', 'title')
        spreadsheet = self.gc.open(spreadsheet_title)
        worksheet = spreadsheet.worksheet(worksheet_title)

        self.assertEqual(worksheet.id,
                         conf.get('Worksheet Properties', 'id'))
        self.assertEqual(worksheet.title,
                         conf.get('Worksheet Properties', 'title'))
        self.assertEqual(worksheet.row_count,
                         conf.getint('Worksheet Properties', 'row_count'))
        self.assertEqual(worksheet.col_count,
                         conf.getint('Worksheet Properties', 'col_count'))


class WorksheetTest(GspreadTest):
    """Test for gspread.Worksheet."""

    @classmethod
    def setUpClass(cls):
        super(WorksheetTest, cls).setUpClass()
        title = cls.config.get('Spreadsheet', 'title')
        cls.spreadsheet = cls.gc.open(title)

    def setUp(self):
        # Set up a clean new worksheet for every test
        self.test_sheet = _create_worksheet(self.spreadsheet)

    def tearDown(self):
        # And delete it afterwards
        self.spreadsheet.del_worksheet(self.test_sheet)

    def test_get_int_addr(self):
        self.assertEqual(self.test_sheet.get_int_addr('ABC3'), (3, 731))

    def test_get_addr_int(self):
        self.assertEqual(self.test_sheet.get_addr_int(3, 731), 'ABC3')
        self.assertEqual(self.test_sheet.get_addr_int(1, 104), 'CZ1')

    def test_addr_converters(self):
        for row in range(1, 257):
            for col in range(1, 512):
                addr = self.test_sheet.get_addr_int(row, col)
                (r, c) = self.test_sheet.get_int_addr(addr)
                self.assertEqual((row, col), (r, c))

    def test_acell(self):
        cell = self.test_sheet.acell('A1')
        self.assertTrue(isinstance(cell, gspread.Cell))

    def test_cell(self):
        cell = self.test_sheet.cell(1, 1)
        self.assertTrue(isinstance(cell, gspread.Cell))

    def test_range(self):
        cell_range = self.test_sheet.range('A1:A5')
        for c in cell_range:
            self.assertTrue(isinstance(c, gspread.Cell))

    def test_update_acell(self):
        value = hashlib.md5(str(time.time())).hexdigest()
        self.test_sheet.update_acell('A2', value)
        self.assertEqual(self.test_sheet.acell('A2').value, value)

    def test_update_cell(self):
        value = hashlib.md5(str(time.time())).hexdigest()
        self.test_sheet.update_cell(1, 2, value)
        self.assertEqual(self.test_sheet.cell(1, 2).value, value)

        self.test_sheet.update_cell(1, 2, 42)
        self.assertEqual(self.test_sheet.cell(1, 2).value, '42')

        self.test_sheet.update_cell(1, 2, 42)
        self.assertEqual(self.test_sheet.cell(1, 2).value, '42')

        self.test_sheet.update_cell(1, 2, 42.01)
        self.assertEqual(self.test_sheet.cell(1, 2).value, '42.01')

        self.test_sheet.update_cell(1, 2, u'Артур')
        self.assertEqual(self.test_sheet.cell(1, 2).value, u'Артур')

    def test_update_cell_multiline(self):
        value = hashlib.md5(str(time.time())).hexdigest()
        value = "%s\n%s" % (value, value)
        self.test_sheet.update_cell(1, 2, value)
        self.assertEqual(self.test_sheet.cell(1, 2).value, value)

    def test_update_cells(self):
        list_len = 10
        value_list = [hashlib.md5(str(time.time() + i)).hexdigest()
                      for i in range(list_len)]
        # Test multiline
        value_list[0] = "%s\n%s" % (value_list[0], value_list[0])

        range_label = 'A1:A%s' % list_len
        cell_list = self.test_sheet.range(range_label)

        for c, v in zip(cell_list, value_list):
            c.value = v

        self.test_sheet.update_cells(cell_list)

        cell_list = self.test_sheet.range(range_label)

        for c, v in zip(cell_list, value_list):
            self.assertEqual(c.value, v)

    def test_resize(self):
        add_num = 10

        new_rows = self.test_sheet.row_count + add_num
        self.test_sheet.add_rows(add_num)
        self.assertEqual(self.test_sheet.row_count, new_rows)

        new_cols = self.test_sheet.col_count + add_num
        self.test_sheet.add_cols(add_num)
        self.assertEqual(self.test_sheet.col_count, new_cols)

        new_rows -= add_num
        new_cols -= add_num
        self.test_sheet.resize(new_rows, new_cols)

        self.assertEqual(self.test_sheet.row_count, new_rows)
        self.assertEqual(self.test_sheet.col_count, new_cols)

    def test_find(self):
        sheet = self.test_sheet
        value = hashlib.md5(str(time.time())).hexdigest()

        sheet.update_cell(2, 10, value)
        sheet.update_cell(2, 11, value)

        cell = sheet.find(value)
        self.assertEqual(cell.value, value)

        value2 = hashlib.md5(str(time.time())).hexdigest()
        value = "%so_O%s" % (value, value2)
        sheet.update_cell(2, 11, value)

        o_O_re = re.compile('[a-z]_[A-Z]%s' % value2)

        cell = sheet.find(o_O_re)
        self.assertEqual(cell.value, value)

    def test_findall(self):
        list_len = 10
        range_label = 'A1:A%s' % list_len
        cell_list = self.test_sheet.range(range_label)
        value = hashlib.md5(str(time.time())).hexdigest()

        for c in cell_list:
            c.value = value
        self.test_sheet.update_cells(cell_list)

        result_list = self.test_sheet.findall(value)

        self.assertEqual(list_len, len(result_list))

        for c in result_list:
            self.assertEqual(c.value, value)

        cell_list = self.test_sheet.range(range_label)

        value = hashlib.md5(str(time.time())).hexdigest()
        for c in cell_list:
            char = chr(random.randrange(ord('a'), ord('z')))
            c.value = "%s%s_%s%s" % (c.value, char, char.upper(), value)

        self.test_sheet.update_cells(cell_list)

        o_O_re = re.compile('[a-z]_[A-Z]%s' % value)

        result_list = self.test_sheet.findall(o_O_re)

        self.assertEqual(list_len, len(result_list))

    def test_get_all_values(self):
        # put in new values, made from three lists
        rows = [["A1", "B1", "", "D1"],
                ["", "b2", "", ""],
                ["", "", "", ""],
                ["A4", "B4", "", "D4"]]
        cell_list = self.test_sheet.range('A1:D1')
        cell_list.extend(self.test_sheet.range('A2:D2'))
        cell_list.extend(self.test_sheet.range('A3:D3'))
        cell_list.extend(self.test_sheet.range('A4:D4'))
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.test_sheet.update_cells(cell_list)

        # read values with get_all_values, get a list of lists
        read_data = self.test_sheet.get_all_values()

        # values should match with original lists
        self.assertEqual(read_data, rows)

    def test_get_all_records(self):
        # put in new values, made from three lists
        rows = [["A1", "B1", "", "D1"],
                [1, "b2", 1.45, ""],
                ["", "", "", ""],
                ["A4", 0.4, "", 4]]
        cell_list = self.test_sheet.range('A1:D1')
        cell_list.extend(self.test_sheet.range('A2:D2'))
        cell_list.extend(self.test_sheet.range('A3:D3'))
        cell_list.extend(self.test_sheet.range('A4:D4'))
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.test_sheet.update_cells(cell_list)

        # first, read empty strings to empty strings
        read_records = self.test_sheet.get_all_records()
        d0 = dict(zip(rows[0], rows[1]))
        d1 = dict(zip(rows[0], rows[2]))
        d2 = dict(zip(rows[0], rows[3]))
        self.assertEqual(read_records[0], d0)
        self.assertEqual(read_records[1], d1)
        self.assertEqual(read_records[2], d2)

        # then, read empty strings to zeros
        read_records = self.test_sheet.get_all_records(empty2zero=True)
        d1 = dict(zip(rows[0], (0, 0, 0, 0)))
        self.assertEqual(read_records[1], d1)

    def test_get_all_records_different_header(self):
        # put in new values, made from three lists
        rows = [["", "", "", ""],
                ["", "", "", ""],
                ["A1", "B1", "", "D1"],
                [1, "b2", 1.45, ""],
                ["", "", "", ""],
                ["A4", 0.4, "", 4]]
        cell_list = self.test_sheet.range('A1:D1')
        cell_list.extend(self.test_sheet.range('A2:D2'))
        cell_list.extend(self.test_sheet.range('A3:D3'))
        cell_list.extend(self.test_sheet.range('A4:D4'))
        cell_list.extend(self.test_sheet.range('A5:D5'))
        cell_list.extend(self.test_sheet.range('A6:D6'))
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.test_sheet.update_cells(cell_list)

        # first, read empty strings to empty strings
        read_records = self.test_sheet.get_all_records(head=3)
        d0 = dict(zip(rows[2], rows[3]))
        d1 = dict(zip(rows[2], rows[4]))
        d2 = dict(zip(rows[2], rows[5]))
        self.assertEqual(read_records[0], d0)
        self.assertEqual(read_records[1], d1)
        self.assertEqual(read_records[2], d2)

        # then, read empty strings to zeros
        read_records = self.test_sheet.get_all_records(empty2zero=True, head=3)
        d1 = dict(zip(rows[2], (0, 0, 0, 0)))
        self.assertEqual(read_records[1], d1)

    def test_append_row(self):
        num_rows = self.test_sheet.row_count
        num_cols = self.test_sheet.col_count
        values = ['o_0'] * (num_cols + 4)
        self.test_sheet.append_row(values)
        self.assertEqual(self.test_sheet.row_count, num_rows + 1)
        self.assertEqual(self.test_sheet.col_count, num_cols + 4)
        read_values = self.test_sheet.row_values(self.test_sheet.row_count)
        self.assertEqual(values, read_values)

    def test_insert_row(self):
        num_rows = self.test_sheet.row_count
        num_cols = self.test_sheet.col_count
        values = ['o_0'] * (num_cols + 4)
        self.test_sheet.insert_row(values, 1)
        self.assertEqual(self.test_sheet.row_count, num_rows + 1)
        self.assertEqual(self.test_sheet.col_count, num_cols + 4)
        read_values = self.test_sheet.row_values(1)
        self.assertEqual(values, read_values)

    def test_export(self):
        list_len = self.test_sheet.row_count

        value_list = [hashlib.md5(str(time.time() + i)).hexdigest()
                      for i in range(list_len)]

        range_label = 'A1:A%s' % list_len
        cell_list = self.test_sheet.range(range_label)

        for c, v in zip(cell_list, value_list):
            c.value = v

        self.test_sheet.update_cells(cell_list)

        exported_data = self.test_sheet.export(format='csv').read()

        csv_value = '\n'.join(value_list)

        self.assertEqual(exported_data, csv_value)


class CellTest(GspreadTest):
    """Test for gspread.Cell."""

    @classmethod
    def setUpClass(cls):
        super(CellTest, cls).setUpClass()
        title = cls.config.get('Spreadsheet', 'title')
        cls.spreadsheet = cls.gc.open(title)

    def setUp(self):
        # Set up a clean new worksheet for every test
        self.test_sheet = _create_worksheet(self.spreadsheet)

    def tearDown(self):
        # And delete it afterwards
        self.spreadsheet.del_worksheet(self.test_sheet)

    def test_properties(self):
        update_value = hashlib.md5(str(time.time())).hexdigest()
        self.test_sheet.update_acell('A1', update_value)
        cell = self.test_sheet.acell('A1')
        self.assertEqual(cell.value, update_value)
        self.assertEqual(cell.row, 1)
        self.assertEqual(cell.col, 1)

    def test_numeric_value(self):
        numeric_value = 1.0 / 1024
        # Use a formula here to avoid issues with differing decimal marks:
        self.test_sheet.update_acell('A1', '= 1 / 1024')
        cell = self.test_sheet.acell('A1')
        self.assertEqual(cell.numeric_value, numeric_value)
        self.assertIsInstance(cell.numeric_value, float)
        self.test_sheet.update_acell('A1', 'Non-numeric value')
        cell = self.test_sheet.acell('A1')
        self.assertIs(cell.numeric_value, None)
