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

  worksheet.set_column(COLS("B:D"), 5, nullptr);
  worksheet.set_column(COLS("F:F"), 8, nullptr, options);
  worksheet.set_column(COLS("H:H"), xwpp::DEF_COL_WIDTH, bold);
  worksheet.set_column(COLS("J:J"), 2, nullptr);
  worksheet.set_column(COLS("L:L"), xwpp::DEF_COL_WIDTH, nullptr, options);

  workbook.save("test_row_col_format13.xlsx");
}
