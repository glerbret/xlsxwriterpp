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
    {.str = "a"},
    {.format = bold, .str = "bc"},
    {.str = "defg"}
  };
  const std::vector<xwpp::rich_string_tuple_t> rich_strings2{
    {.str = "a"},
    {.format = bold, .str = "bcdef"},
    {.str = "g"}
  };
  const std::vector<xwpp::rich_string_tuple_t> rich_strings3{
    {.str = "abc"},
    {.format = italic, .str = "de"},
    {.str = "fg"}
  };
  const std::vector<xwpp::rich_string_tuple_t> rich_strings4{
    {.format = italic, .str = "abcd"},
    {.str = "efg"}
  };

  worksheet.write_rich_string("A3", rich_strings1);
  worksheet.write_rich_string("B4", rich_strings3);
  worksheet.write_rich_string("C5", rich_strings1);
  worksheet.write_rich_string("D6", rich_strings3);
  worksheet.write_rich_string("E7", rich_strings2);
  worksheet.write_rich_string("F8", rich_strings4);

  workbook.save("test_rich_string07.xlsx");
}
