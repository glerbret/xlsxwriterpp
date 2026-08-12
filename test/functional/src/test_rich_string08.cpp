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

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  xwpp::format_t* centered = workbook.add_format();
  centered->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);

  worksheet.write_string(CELL("A1"), "Foo", bold);
  worksheet.write_string(CELL("A2"), "Bar", italic);

  const std::vector<xwpp::rich_string_tuple_t> rich_strings{
    {.str_ = "ab"},
    {.format_ = bold, .str_ = "cd"},
    {.str_ = "efg"}
  };
  worksheet.write_rich_string(CELL("A3"), rich_strings, centered);

  workbook.save("test_rich_string08.xlsx");
}
