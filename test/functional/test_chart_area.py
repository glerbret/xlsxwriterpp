# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with Xlsxwriter++ against a file created by Excel.

    """

    def test_chart_area01(self):
        self.run_exe_test('test_chart_area01')

    def test_chart_area02(self):
        self.run_exe_test('test_chart_area02')

    def test_chart_area03(self):
        self.run_exe_test('test_chart_area03')

    def test_chart_area05(self):
        self.run_exe_test('test_chart_area05')

    def test_chart_area06(self):
        self.run_exe_test('test_chart_area06')
