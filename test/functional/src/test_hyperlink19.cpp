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

  // This test requires that we check if the cell that the hyperlink refers
  // to is a string. If it isn't be have to add a display attribute. However,
  // looking up the cell is currently too expensive.
  worksheet.write_url(CELL("A1"), "http://www.perl.com/");
  worksheet.write_formula_num(CELL("A1"), "=1+1", nullptr, 2);

  workbook.save("test_hyperlink19.xlsx");
}
