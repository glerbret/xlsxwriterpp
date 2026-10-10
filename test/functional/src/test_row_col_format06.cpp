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

  const xwpp::format_t* bold   = workbook.format_builder().bold().build();
  const xwpp::format_t* italic = workbook.format_builder().italic().build();

  worksheet.set_column(0, 0, 8.43, bold);
  worksheet.set_column(2, 2, 8.43, italic);

  worksheet.write(0, 0, "Foo");
  worksheet.write(0, 2, "Bar");

  workbook.save("test_row_col_format06.xlsx");
}
