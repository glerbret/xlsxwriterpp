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

  const xwpp::image_options_t options1{.url_ = "https://github.com/jmcnamara"};
  const xwpp::image_options_t options2{.url_ = "https://github.com"};

  worksheet.insert_image(CELL("A1"), "images/blue.png");
  worksheet.insert_image(CELL("B3"), "images/red.jpg", options1);
  worksheet.insert_image(CELL("D5"), "images/yellow.jpg");
  worksheet.insert_image(CELL("F9"), "images/grey.png", options2);

  workbook.save("test_hyperlink34.xlsx");
}
