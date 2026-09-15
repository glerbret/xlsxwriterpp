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

  worksheet.set_column(0, 0, 32);

  // Test floating point formatting.
  worksheet.write(0, 0, 0);
  worksheet.write(1, 0, 1);
  worksheet.write(2, 0, -1);
  worksheet.write(3, 0, 1.2);
  worksheet.write(4, 0, -1.2);
  worksheet.write(5, 0, 1.2E8);
  worksheet.write(6, 0, 1.2E+20);
  worksheet.write(7, 0, 1.2E-20);
  worksheet.write(8, 0, -1.2E+20);
  worksheet.write(9, 0, -1.2E-20);
  worksheet.write(10, 0, 1.E+100);
  worksheet.write(11, 0, 1.E-100);

  workbook.save("test_types11.xlsx");
}
