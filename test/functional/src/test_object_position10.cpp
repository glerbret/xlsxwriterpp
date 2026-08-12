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

  const xwpp::image_options_t options{.object_position_ = xwpp::object_position_t::MOVE_AND_SIZE_AFTER};
  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  const xwpp::row_col_options_t col_hidden{.hidden_ = true};
  worksheet.set_column(COLS("E:E"), xwpp::DEF_COL_WIDTH, nullptr, col_hidden);

  workbook.save("test_object_position10.xlsx");
}
