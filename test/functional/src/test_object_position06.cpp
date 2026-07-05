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

  xwpp::image_options_t options = {.object_position_ = xwpp::object_position_t::MOVE_AND_SIZE_AFTER};
  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  xwpp::row_col_options_t hidden = {.hidden_ = true};
  worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);

  workbook.save("test_object_position06.xlsx");
}
