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
  xwpp::format_t* format1      = workbook.add_format();
  xwpp::format_t* format2      = workbook.add_format();
  xwpp::format_t* format3      = workbook.add_format();

  workbook.unset_default_url_format();

  format1->set_hyperlink();

  format2->set_underline(xwpp::format_underlines_t::SINGLE);
  format2->set_font_color(xwpp::color_t::RED);

  format3->set_font_color(xwpp::color_t::BLUE);
  format3->set_underline(xwpp::format_underlines_t::SINGLE);

  worksheet.write_url(CELL("A1"), "http://www.python.org/1", format1);
  worksheet.write_url(CELL("A2"), "http://www.python.org/2", format2);
  worksheet.write_url(CELL("A3"), "http://www.python.org/3", format3);

  workbook.save("test_hyperlink30.xlsx");
}
