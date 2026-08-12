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

  worksheet.write_number(CELL("A1"), 1234);
  worksheet.write_number(CELL("C7"), 1234);
  worksheet.write_number(CELL("G14"), 1234);

  worksheet.write_comment(CELL("A1"), "Some text");
  worksheet.write_comment(CELL("D1"), "Some text");
  worksheet.write_comment(CELL("C7"), "Some text");
  worksheet.write_comment(CELL("E10"), "Some text");
  worksheet.write_comment(CELL("G14"), "Some text");

  // Repeat above to check for overwrite leaks.
  worksheet.write_string(CELL("A1"), "Foo");
  worksheet.write_string(CELL("C7"), "Bar");
  worksheet.write_string(CELL("G14"), "Baz");

  worksheet.write_comment(CELL("A1"), "Some text");
  worksheet.write_comment(CELL("D1"), "Some text");
  worksheet.write_comment(CELL("C7"), "Some text");
  worksheet.write_comment(CELL("E10"), "Some text");
  worksheet.write_comment(CELL("G14"), "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment56.xlsx");
}
