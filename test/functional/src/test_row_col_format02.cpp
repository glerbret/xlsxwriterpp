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

  const xwpp::format_t* bold = workbook.format_builder().bold().build();

  worksheet.set_row(0, 15, bold);

  worksheet.write(0, 0, "Foo");

  workbook.save("test_row_col_format02.xlsx");
}
