# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with Xlsxwriter++ against a file created by Excel.
    """

    def test_chart_chartarea01(self):
        self.run_exe_test('test_chart_chartarea01')

    # chartarea02 is for a deprecated api in Perl.

    def test_chart_chartarea03(self):
        self.run_exe_test('test_chart_chartarea03')

    # chartarea04 is an unsupported stock chart.

    def test_chart_chartarea05(self):
        self.run_exe_test('test_chart_chartarea05')

    def test_chart_chartarea06(self):
        self.run_exe_test('test_chart_chartarea06')
