/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Tests for the row range.
  worksheet.write_number(0, 0, 123);
  worksheet.write_number(1048575, 0, 456);

  // TODO Throw an exception, add explicit tests
  //  worksheet.write_number(-1, 0, 123);
  //  worksheet.write_number(1048576, 0, 456);

  workbook.save("test_data02.xlsx");
}
