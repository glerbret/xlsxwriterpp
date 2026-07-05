# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with Xlsxwriter++ against a file created by Excel.
    """

    def test_types02(self):
        self.run_exe_test('test_types02')

    def test_types08(self):
        self.run_exe_test('test_types08')

    def test_types11(self):
        self.run_exe_test('test_types11')
