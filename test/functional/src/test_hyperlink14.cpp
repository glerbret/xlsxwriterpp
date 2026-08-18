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
  xwpp::format_t* format       = workbook.add_format();

  format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);

  worksheet.write_string(CELL("A1"), "Perl Home");

  worksheet.merge_range(RANGE("C4:E5"), "http://www.perl.org/", format);
  worksheet.write_url(CELL("C4"), "http://www.perl.org/", format, "Perl Home", "");

  workbook.save("test_hyperlink14.xlsx");
}
