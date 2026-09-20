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

  worksheet.set_column("C:H", 10.288);

  worksheet.add_table("C3:F13");
  worksheet.add_table("F15:H20");
  worksheet.add_table("C23:D30");

  // Add other objects to check rId handling.
  workbook.unset_default_url_format();
  worksheet.write_url("A1", "http://perl.com/");
  worksheet.write_url("C1", "http://perl.com/");

  worksheet.set_comments_author("John");
  worksheet.write_comment("H1", "Test1");
  worksheet.write_comment("J1", "Test2");

  worksheet.insert_image("A4", "images/blue.png");

  workbook.save("test_table06.xlsx");
}
