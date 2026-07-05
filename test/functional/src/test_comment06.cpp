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

  xwpp::comment_options_t options = {.visible_ = xwpp::comment_display_t::VISIBLE};

  worksheet.write_comment(CELL("A1"), "Some text");
  worksheet.write_comment(CELL("A2"), "Some text");

  worksheet.write_comment(CELL("A3"), "Some text", options);

  worksheet.write_comment(CELL("A4"), "Some text");
  worksheet.write_comment(CELL("A5"), "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment06.xlsx");
}
