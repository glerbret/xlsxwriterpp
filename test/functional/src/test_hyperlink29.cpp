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

  const xwpp::format_t* format1 = workbook.format_builder().hyperlink().build();
  const xwpp::format_t* format2 =
    workbook.format_builder().underline(xwpp::format_underlines_t::SINGLE).font_color(xwpp::color_t::red()).build();

  worksheet.write_url("A1", "http://www.perl.org/", format1);
  worksheet.write_url("A2", "http://www.perl.com/", format2);

  workbook.save("test_hyperlink29.xlsx");
}
