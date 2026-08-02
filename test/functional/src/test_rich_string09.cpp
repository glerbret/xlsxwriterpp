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

  std::vector<xwpp::rich_string_tuple_t> rich_strings1 = {
    {.str_ = "a"},
    {.format_ = bold, .str_ = "bc"},
    {.str_ = "defg"}
  };
  worksheet.write_rich_string(CELL("A3"), rich_strings1, nullptr);

  // The following use cases fail validation and should be ignored.
  // TODO Throw exception (and don't ignore) on xlsxwriter++, test to add
  /*
    // Empty string.
    std::vector<xwpp::rich_string_tuple_t> rich_strings3 = {
        {.str_ = ""},
        {.format_ = bold, .str_ = "bc"},
        {.str_ = "defg"}
    };
    worksheet.write_rich_string(CELL("A3"), rich_strings3, nullptr);

    // Empty array.
    std::vector<xwpp::rich_string_tuple_t> rich_strings4;
    worksheet.write_rich_string(CELL("A3"), rich_strings4, nullptr);

    // Single string.
    std::vector<xwpp::rich_string_tuple_t> rich_strings5 = {
        {.format_ = bold, .str_ = "foo"}
    };
    worksheet.write_rich_string(CELL("A3"), rich_strings5, nullptr);
  */
  workbook.save("test_rich_string09.xlsx");
}
