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

  worksheet.write_dynamic_formula(CELL("B1"), "=_xlfn.UNIQUE(A1)");
  worksheet.write_number(CELL("A1"), 0);

  workbook.save("test_dynamic_array02.xlsx");
}
