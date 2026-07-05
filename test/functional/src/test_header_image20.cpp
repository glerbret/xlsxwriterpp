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

  xwpp::header_footer_options_t header_options = {.image_center_ = "images/watermark.png"};

  worksheet.set_header("&C&G", header_options);

  workbook.save("test_header_image20.xlsx");
}
