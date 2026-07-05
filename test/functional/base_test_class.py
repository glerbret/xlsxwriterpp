# Base test class for Xlsxwriter++ functional tests.
#
# Copyright 2026, Grégory Lerbret
#
# Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).


import unittest
import os
import sys
from helper_functions import _compare_xlsx_files


class XLSXBaseTest(unittest.TestCase):
    """
    Test file created with Xlsxwriter++ against a file created by Excel.
    """

    def setUp(self):
        self.max_diff = None

        self.no_system_error = 0
        self.got_filename = ''
        self.exp_filename = ''
        self.ignore_files = []
        self.ignore_elements = {}

    def run_exe_test(self, exe_name, exp_filename=None, must_throw=False):
        """Run C++ exe and compare output xlsx file with the Excel file."""

        # Create the executable command
        if sys.platform == 'win32':
            command = f"cd test\\functional\\src && {exe_name}.exe"
        else:
            command = f"cd test/functional/src && ./{exe_name}"

        # Run the C++ executable to generate the "got" xlsx file.
        got = os.system(command)

        # C++ executable is expected to throw
        if must_throw is True:
            self.assertNotEqual(got, self.no_system_error)
            return

        self.assertEqual(got, self.no_system_error)

        # Create the path/file names for the xlsx/xlsm files to compare.
        if exp_filename and exp_filename.endswith('.xlsm'):
            got_filename = exe_name.replace('test_', '') + '.xlsm'
        else:
            got_filename = exe_name.replace('test_', '') + '.xlsx'

        if not exp_filename:
            exp_filename = got_filename

        self.got_filename = 'test/functional/src/test_' + got_filename
        self.exp_filename = 'test/functional/xlsx_files/' + exp_filename

        # Do the comparison between the files.
        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)
        self.assertEqual(exp, got)

    def tearDown(self):
        # Cleanup.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)

        self.ignore_files = []
        self.ignore_elements = {}
