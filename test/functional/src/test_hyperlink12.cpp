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

  format->set_underline(xwpp::format_underlines_t::SINGLE);
  format->set_font_color(xwpp::color_t::BLUE);

  worksheet.write_url(CELL("A1"), "mailto:jmcnamara@cpan.org", format);
  worksheet.write_url(CELL("A3"), "ftp://perl.org/", format);

  workbook.save("test_hyperlink12.xlsx");
}
