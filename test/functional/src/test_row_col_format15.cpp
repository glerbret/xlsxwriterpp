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

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  worksheet.set_column(16383, 16383, 8.43, italic);

  workbook.save("test_row_col_format15.xlsx");
}
