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

  xwpp::header_footer_options_t header_options = {.image_left_ = "images/red.jpg"};
  xwpp::header_footer_options_t footer_options = {.image_left_ = "images/blue.jpg"};

  worksheet.set_header("&L&G", header_options);
  worksheet.set_footer("&L&G", footer_options);

  workbook.save("test_header_image05.xlsx");
}
