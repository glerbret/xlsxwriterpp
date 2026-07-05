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

  worksheet.write_string(CELL("A1"), "123");
  worksheet.write_string(CELL("C3"), "123");
  worksheet.write_string(CELL("E5"), "123");

  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "A1 C3 E5");

  workbook.save("test_ignore_error04.xlsx");
}
