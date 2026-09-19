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

  worksheet.set_column("A:A", 100);
  worksheet.set_column("F:H", 8);
  worksheet.set_column("C:D", 12);
  worksheet.set_column("A:A", 10);
  worksheet.set_column("XFD:XFD", 5);
  worksheet.set_column("ZZ:ZZ", 3);

  workbook.save("test_set_column09.xlsx");
}
