# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_chart_high_low_lines01(self):
        self.run_exe_test('test_chart_high_low_lines01')

    def test_chart_high_low_lines02(self):
        self.run_exe_test('test_chart_high_low_lines02')
