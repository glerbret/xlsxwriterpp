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

  xwpp::header_footer_options_t header_options = {.image_left_   = "images/black_72.jpg",
                                                  .image_center_ = "images/black_150.jpg",
                                                  .image_right_  = "images/black_300.jpg"};

  worksheet.set_header("&L&G&C&G&R&G", header_options);

  workbook.save("test_header_image13.xlsx");
}
