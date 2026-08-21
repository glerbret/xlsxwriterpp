/*
 * An example of using the Xlsxwriter++ library to write some "rich strings",
 * i.e., strings with multiple formats.
 *
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

  // Set up some formats to use.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  xwpp::format_t* red = workbook.add_format();
  red->set_font_color(xwpp::color_t::RED);

  xwpp::format_t* blue = workbook.add_format();
  blue->set_font_color(xwpp::color_t::BLUE);

  xwpp::format_t* center = workbook.add_format();
  center->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);

  xwpp::format_t* superscript = workbook.add_format();
  superscript->set_font_script(xwpp::format_scripts_t::SUPERSCRIPT);

  // Make the first column wider for clarity.
  worksheet.set_column(0, 0, 30);

  // Create and write some rich strings with multiple formats.

  // Example 1. Some bold and italic in the same string.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str_ = "This is "},
      {.format_ = bold, .str_ = "bold"},
      {.str_ = " and this is "},
      {.format_ = italic, .str_ = "italic"},
    };

    worksheet.write_rich_string(CELL("A1"), rich_string, nullptr);
  }

  // Example 2. Some red and blue coloring in the same string.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str_ = "This is "},
      {.format_ = red, .str_ = "red"},
      {.str_ = " and this is "},
      {.format_ = blue, .str_ = "blue"},
    };

    worksheet.write_rich_string(CELL("A3"), rich_string, nullptr);
  }

  // Example 3. A rich string plus cell formatting.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str_ = "Some "},
      {.format_ = bold, .str_ = "bold text"},
      {.str_ = " centered"},
    };

    // Note that this example also has a "center" cell format.
    worksheet.write_rich_string(CELL("A5"), rich_string, center);
  }

  // Example 4. A math example with a superscript.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.format_ = italic,      .str_ = "j =k" },
      {.format_ = superscript, .str_ = "(n-1)"},
    };

    worksheet.write_rich_string(CELL("A7"), rich_string, center);
  }

  workbook.save("rich_strings.xlsx");
}
