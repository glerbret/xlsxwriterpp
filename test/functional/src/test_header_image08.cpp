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

  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_comment(CELL("B2"), "Some text");

  worksheet.set_comments_author("John");

  const xwpp::header_footer_options_t header_options{.image_left_ = "images/red.jpg"};

  worksheet.set_header("&L&G", header_options);

  workbook.save("test_header_image08.xlsx");
}
