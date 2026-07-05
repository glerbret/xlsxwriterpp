# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_properties01(self):
        self.run_exe_test('test_properties01')

    def test_properties02(self):
        self.run_exe_test('test_properties02')

    def test_properties03(self):
        self.run_exe_test('test_properties03')

    def test_properties04(self):
        self.run_exe_test('test_properties04')

    def test_properties05(self):
        self.run_exe_test('test_properties05')
