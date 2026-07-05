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

  worksheet.write_formula(CELL("A1"), "=B1");
  worksheet.write_formula(CELL("A2"), "=B1");
  worksheet.write_formula(CELL("A3"), "=B3");

  worksheet.ignore_errors(xwpp::ignore_errors_t::FORMULA_DIFFERS, "A2");

  workbook.save("test_ignore_error06.xlsx");
}
