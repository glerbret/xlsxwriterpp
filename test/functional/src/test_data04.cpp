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

  worksheet.write_string(0, 0, "Foo");
  worksheet.write_string(0, 1, "Bar");
  worksheet.write_string(1, 0, "Bing");
  worksheet.write_string(2, 0, "Buzz");
  worksheet.write_string(1048575, 0, "End");

  workbook.save("test_data04.xlsx");
}
