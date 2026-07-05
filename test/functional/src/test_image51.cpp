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

  xwpp::image_options_t options1 = {.url_ = "https://duckduckgo.com/?q=1"};
  xwpp::image_options_t options2 = {.url_ = "https://duckduckgo.com/?q=2"};

  worksheet.insert_image(CELL("E9"), "images/red.png", options1);
  worksheet.insert_image(CELL("E13"), "images/red2.png", options2);

  workbook.save("test_image51.xlsx");
}
