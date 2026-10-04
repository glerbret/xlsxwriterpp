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

  const xwpp::format_t* bold   = workbook.format_builder().bold().build();
  const xwpp::format_t* italic = workbook.format_builder().italic().build();

  worksheet.write("A1", "Foo", bold);
  worksheet.write("A2", "Bar", italic);

  const std::vector<xwpp::rich_string_tuple_t> rich_strings1{
    {.str_ = "a"},
    {.format_ = bold, .str_ = "bc"},
    {.str_ = "defg"}
  };
  worksheet.write_rich_string("A3", rich_strings1);

  workbook.save("test_rich_string09.xlsx");
}
