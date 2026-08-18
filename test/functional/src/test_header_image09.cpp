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

  worksheet1.write_string(CELL("A1"), "Foo");
  worksheet1.write_comment(CELL("B2"), "Some text");

  worksheet1.set_comments_author("John");

  const xwpp::header_footer_options_t header_options{.image_left_ = "images/red.jpg"};

  worksheet2.set_header("&L&G", header_options);

  workbook.save("test_header_image09.xlsx");
}
