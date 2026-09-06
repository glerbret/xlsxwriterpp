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

  worksheet.set_default_row(24, true);

  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_string(CELL("A10"), "Bar");
  worksheet.write_string(CELL("A20"), "Baz");

  for(xwpp::row_num_t row_num{1}; row_num <= 8; row_num++)
  {
    worksheet.set_row(row_num, 24);
  }

  for(xwpp::row_num_t row_num{10}; row_num <= 19; row_num++)
  {
    worksheet.set_row(row_num, 24);
  }

  workbook.save("test_default_row05.xlsx");
}
