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

  xwpp::format_t* red_format = workbook.add_format();

  red_format->set_font_color(xwpp::color_t::RED);

  worksheet.write_string(CELL("A1"), "Foo", red_format);
  worksheet.write_string(CELL("A2"), "Bar", nullptr);

  std::vector<xwpp::rich_string_tuple_t> rich_strings = {
    {.str_ = "ab"},
    {.format_ = red_format, .str_ = "cde"},
    {.str_ = "fg"}
  };
  worksheet.write_rich_string(CELL("A3"), rich_strings, nullptr);

  workbook.save("test_rich_string06.xlsx");
}
