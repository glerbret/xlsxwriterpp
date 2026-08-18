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

  const xwpp::row_col_options_t col_options{.hidden_ = true};
  worksheet.set_column(COLS("B:B"), 5, nullptr, col_options);

  worksheet.insert_image(CELL("E9"), "images/red.png");

  workbook.save("test_object_position14.xlsx");
}
