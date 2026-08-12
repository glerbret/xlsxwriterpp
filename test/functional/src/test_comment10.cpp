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

  const xwpp::comment_options_t options{.color_ = xwpp::color_t(0x98FE97)};

  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_comment(CELL("B2"), "Some text", options);

  worksheet.set_comments_author("John");

  workbook.save("test_comment10.xlsx");
}
