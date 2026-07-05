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
  worksheet.write_formula(CELL("A2"), "=1/0");

  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "A1");
  worksheet.ignore_errors(xwpp::ignore_errors_t::EVAL_ERROR, "A2");

  workbook.save("test_ignore_error05.xlsx");
}
