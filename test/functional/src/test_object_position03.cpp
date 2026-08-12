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

  const xwpp::image_options_t options{.object_position_ = xwpp::object_position_t::DONT_MOVE_DONT_SIZE};
  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  workbook.save("test_object_position03.xlsx");
}
