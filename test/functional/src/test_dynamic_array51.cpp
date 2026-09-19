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

  worksheet.write_dynamic_array_formula("A1:A1", "=AVERAGE(TIMEVALUE(B1:B2))", nullptr, 0);
  worksheet.write("B1", "12:00");
  worksheet.write("B2", "12:00");

  workbook.save("test_dynamic_array51.xlsx");
}
