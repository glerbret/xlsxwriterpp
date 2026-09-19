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
  worksheet.write_url("A3", "http://www.perl.org/", "Perl home");
  worksheet.write_url("A5", "http://www.perl.org/", "Perl home", "Tool Tip");
  worksheet.write_url("A7", "http://www.cpan.org/", "CPAN", "Download");

  workbook.save("test_hyperlink05.xlsx");
}
