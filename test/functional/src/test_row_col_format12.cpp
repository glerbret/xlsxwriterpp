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

  worksheet.set_column(2, 2, xwpp::DEF_COL_WIDTH, nullptr, options);

  workbook.save("test_row_col_format12.xlsx");
}
