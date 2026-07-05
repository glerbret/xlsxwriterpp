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

  xwpp::image_options_t image_options = {.y_offset_        = 4,
                                         .object_position_ = xwpp::object_position_t::MOVE_AND_SIZE_AFTER};
  worksheet.insert_image(CELL("E9"), "images/red.png", image_options);

  xwpp::row_col_options_t row_options = {.hidden_ = true};
  worksheet.set_row(8, 30, nullptr, row_options);

  workbook.save("test_image46.xlsx");
}
