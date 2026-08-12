/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  const xwpp::header_footer_options_t header_options1{.image_left_ = "images/red.jpg"};
  const xwpp::header_footer_options_t header_options2{.image_right_ = "images/red.jpg"};

  worksheet1.set_header("&L&G", header_options1);
  worksheet2.set_header("&L&G", header_options1);

  worksheet1.set_footer("&R&G", header_options2);
  worksheet2.set_footer("&R&G", header_options2);

  workbook.save("test_header_image16.xlsx");
}
