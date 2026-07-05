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

  worksheet.write_dynamic_formula_num(CELL("A1"), "=AVERAGE(TIMEVALUE(B1:B2))", nullptr, 0);
  worksheet.write_string(CELL("B1"), "12:00");
  worksheet.write_string(CELL("B2"), "12:00");

  workbook.save("test_dynamic_array53.xlsx");
}
