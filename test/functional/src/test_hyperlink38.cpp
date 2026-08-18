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

  const xwpp::image_options_t options{.url_ = "internal:Sheet1!A1"};

  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  workbook.save("test_hyperlink38.xlsx");
}
