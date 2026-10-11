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

  const xwpp::header_footer_options_t header_options{
    .image_left   = "images/red.jpg",
    .image_center = "images/blue.jpg",
    .image_right  = "images/yellow.jpg",
  };
  worksheet.set_header("&L&G&C&G&R&G", header_options);

  workbook.save("test_header_image53.xlsx");
}
