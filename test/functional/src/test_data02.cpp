/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Tests for the row range.
  worksheet.write(0, 0, 123);
  worksheet.write(1048575, 0, 456);

  workbook.save("test_data02.xlsx");
}
