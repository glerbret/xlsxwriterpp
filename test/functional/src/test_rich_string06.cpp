/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  const xwpp::format_t* red = workbook.format_builder().font_color(xwpp::color_t::red()).build();

  worksheet.write("A1", "Foo", red);
  worksheet.write("A2", "Bar", nullptr);

  const std::vector<xwpp::rich_string_tuple_t> rich_strings{
    {.str_ = "ab"},
    {.format_ = red, .str_ = "cde"},
    {.str_ = "fg"}
  };
  worksheet.write_rich_string("A3", rich_strings);

  workbook.save("test_rich_string06.xlsx");
}
