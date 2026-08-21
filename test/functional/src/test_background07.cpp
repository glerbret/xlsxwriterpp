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

  worksheet.insert_image(CELL("E9"), "images/logo.jpg");
  worksheet.set_background("images/logo.jpg");

  const xwpp::header_footer_options_t header_options{.image_center_ = "images/blue.jpg"};
  worksheet.set_header("&C&G", header_options);

  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_comment(CELL("B2"), "Some text");
  worksheet.set_comments_author("John");

  workbook.save("test_background07.xlsx");
}
