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

  workbook.unset_default_url_format();

  worksheet1.write_url(CELL("A1"), "http://www.perl.org/");
  worksheet1.write_url(CELL("D4"), "http://www.perl.org/");
  worksheet1.write_url(CELL("A8"), "http://www.perl.org/");
  worksheet1.write_url(CELL("B6"), "http://www.cpan.org/");
  worksheet1.write_url(CELL("F12"), "http://www.cpan.org/");

  worksheet2.write_url(CELL("C2"), "http://www.google.com/");
  worksheet2.write_url(CELL("C5"), "http://www.cpan.org/");
  worksheet2.write_url(CELL("C7"), "http://www.perl.org/");

  workbook.save("test_hyperlink03.xlsx");
}
