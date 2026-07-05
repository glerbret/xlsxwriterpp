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

  worksheet.set_column(COLS("C:F"), 10.288);

  worksheet.add_table(RANGE("C2:F6"));

  worksheet.write_string(2, 2, "Foo");
  worksheet.write_string(3, 2, "Bar");
  worksheet.write_string(4, 2, "Baz");

  worksheet.write_number(2, 3, 1234);
  worksheet.write_number(3, 3, 1256);
  worksheet.write_number(4, 3, 2234);

  worksheet.write_number(2, 4, 2000);
  worksheet.write_number(3, 4, 4000);
  worksheet.write_number(4, 4, 3000);

  workbook.save("test_table12.xlsx");
}
