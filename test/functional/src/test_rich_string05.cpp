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

  worksheet.set_column(0, 0, 30);

  worksheet.write("A1", "Foo", bold);
  worksheet.write("A2", "Bar", italic);

  const std::vector<xwpp::rich_string_tuple_t> rich_strings{
    {.str_ = "This is "},
    {.format_ = bold, .str_ = "bold"},
    {.str_ = " and this is "},
    {.format_ = italic, .str_ = "italic"}
  };
  worksheet.write_rich_string("A3", rich_strings);

  workbook.save("test_rich_string05.xlsx");
}
