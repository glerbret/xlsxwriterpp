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

  const xwpp::format_t* format =
    workbook.format_builder().underline(xwpp::format_underlines_t::SINGLE).font_color(xwpp::color_t::blue()).build();

  worksheet.write_url("A1", "http://www.perl.org/", format);

  workbook.save("test_hyperlink11.xlsx");
}
