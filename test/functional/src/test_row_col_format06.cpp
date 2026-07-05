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

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  worksheet.set_column(0, 0, 8.43, bold);
  worksheet.set_column(2, 2, 8.43, italic);

  worksheet.write_string(0, 0, "Foo");
  worksheet.write_string(0, 2, "Bar");

  workbook.save("test_row_col_format06.xlsx");
}
