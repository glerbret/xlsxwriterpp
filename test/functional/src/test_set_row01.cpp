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

  worksheet.set_row(0, 0.75);
  worksheet.set_row(1, 1.50);
  worksheet.set_row(2, 2.25);
  worksheet.set_row(3, 3);

  worksheet.set_row(11, 9);
  worksheet.set_row(12, 9.75);
  worksheet.set_row(13, 10.50);
  worksheet.set_row(14, 11.25);

  worksheet.set_row(18, 14.25);
  worksheet.set_row(20, 15.75);
  worksheet.set_row(21, 16.50);

  workbook.save("test_set_row01.xlsx");
}
