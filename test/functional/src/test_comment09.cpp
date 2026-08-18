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

  const xwpp::comment_options_t options1{.author_ = "John"};
  const xwpp::comment_options_t options2{.author_ = "Perl"};

  worksheet.write_comment(CELL("A1"), "Some text", options1);
  worksheet.write_comment(CELL("A2"), "Some text", options2);
  worksheet.write_comment(CELL("A3"), "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment09.xlsx");
}
