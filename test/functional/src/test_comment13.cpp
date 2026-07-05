/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::comment_options_t options = {.font_name_ = "Courier", .font_size_ = 10, .font_family_ = 3};

  worksheet.write_string(CELL("A1"), "Foo");

  worksheet.write_comment(CELL("B2"), "Some text", options);

  worksheet.set_comments_author("John");

  workbook.save("test_comment13.xlsx");
}
