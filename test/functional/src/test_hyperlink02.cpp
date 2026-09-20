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

  workbook.unset_default_url_format();

  worksheet.write_url("A1", "http://www.perl.org/");
  worksheet.write_url("D4", "http://www.perl.org/");
  worksheet.write_url("A8", "http://www.perl.org/");
  worksheet.write_url("B6", "http://www.cpan.org/");
  worksheet.write_url("F12", "http://www.cpan.org/");

  workbook.save("test_hyperlink02.xlsx");
}
