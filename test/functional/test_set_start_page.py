# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_set_start_page01(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_set_start_page01')

    def test_set_start_page02(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_set_start_page02')

    def test_set_start_page03(self):
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}
        self.run_exe_test('test_set_start_page03')
