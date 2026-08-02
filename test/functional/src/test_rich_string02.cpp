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

  xwpp::format_t* bold   = workbook.add_format();
  xwpp::format_t* italic = workbook.add_format();

  bold->set_bold();
  italic->set_italic();

  worksheet.write_string(CELL("A1"), "Foo", bold);
  worksheet.write_string(CELL("A2"), "Bar", italic);

  std::vector<xwpp::rich_string_tuple_t> rich_strings = {
    {.str_ = "abcd"},
    {.format_ = italic, .str_ = "ef"},
    {.str_ = "g"}
  };
  worksheet.write_rich_string(CELL("A3"), rich_strings, nullptr);

  workbook.save("test_rich_string02.xlsx");
}
