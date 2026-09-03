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

  worksheet.set_default_row(15, true);

  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_string(CELL("A10"), "Bar");

  for(xwpp::row_num_t row{1}; row <= 8; row++)
  {
    worksheet.set_row(row, 15);
  }

  workbook.save("test_default_row02.xlsx");
}
