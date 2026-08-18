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
  worksheet.set_column(COLS("B:B"), xwpp::DEF_COL_WIDTH, nullptr, col_options);

  const xwpp::image_options_t image_options{.x_offset_ = 128};
  worksheet.insert_image(CELL("B9"), "images/red.png", image_options);

  workbook.save("test_object_position19.xlsx");
}
