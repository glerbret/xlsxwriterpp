# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).

import base_test_class


class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_data_validation01(self):
        self.run_exe_test('test_data_validation01')

    def test_data_validation02(self):
        self.run_exe_test('test_data_validation02')

    def test_data_validation03(self):
        self.run_exe_test('test_data_validation03')

    def test_data_validation04(self):
        self.run_exe_test('test_data_validation04', must_throw=True)

    def test_data_validation05(self):
        self.run_exe_test('test_data_validation05', must_throw=True)

    def test_data_validation06(self):
        self.run_exe_test('test_data_validation06', must_throw=True)

    def test_data_validation07(self):
        self.run_exe_test('test_data_validation07')

    def test_data_validation08(self):
        self.run_exe_test('test_data_validation08')
