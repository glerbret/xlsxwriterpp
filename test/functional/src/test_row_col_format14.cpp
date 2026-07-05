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

  xwpp::row_col_options_t options = {.hidden_ = true, .level_ = 0, .collapsed_ = false};

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  worksheet.set_column(1, 3, 5);
  worksheet.set_column(5, 5, 8);
  worksheet.set_column(7, 7, xwpp::DEF_COL_WIDTH, bold);
  worksheet.set_column(9, 9, 2);
  worksheet.set_column(11, 11, xwpp::DEF_COL_WIDTH, nullptr, options);

  workbook.save("test_row_col_format14.xlsx");
}
