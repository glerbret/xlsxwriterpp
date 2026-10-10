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
  const xwpp::format_t* bold   = workbook.format_builder().bold().build();
  const xwpp::format_t* italic = workbook.format_builder().italic().build();
  const xwpp::format_t* red    = workbook.format_builder().font_color(xwpp::color_t::red()).build();
  const xwpp::format_t* blue   = workbook.format_builder().font_color(xwpp::color_t::blue()).build();
  const xwpp::format_t* center = workbook.format_builder().align(xwpp::format_horizontal_alignments_t::CENTER).build();

  const xwpp::format_t* superscript =
    workbook.format_builder().font_script(xwpp::format_scripts_t::SUPERSCRIPT).build();

  // Make the first column wider for clarity.
  worksheet.set_column(0, 0, 30);

  // Create and write some rich strings with multiple formats.

  // Example 1. Some bold and italic in the same string.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str = "This is "},
      {.format = bold, .str = "bold"},
      {.str = " and this is "},
      {.format = italic, .str = "italic"},
    };

    worksheet.write_rich_string("A1", rich_string);
  }

  // Example 2. Some red and blue coloring in the same string.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str = "This is "},
      {.format = red, .str = "red"},
      {.str = " and this is "},
      {.format = blue, .str = "blue"},
    };

    worksheet.write_rich_string("A3", rich_string);
  }

  // Example 3. A rich string plus cell formatting.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.str = "Some "},
      {.format = bold, .str = "bold text"},
      {.str = " centered"},
    };

    // Note that this example also has a "center" cell format.
    worksheet.write_rich_string("A5", rich_string, center);
  }

  // Example 4. A math example with a superscript.
  {
    const std::vector<xwpp::rich_string_tuple_t> rich_string{
      {.format = italic,      .str = "j =k" },
      {.format = superscript, .str = "(n-1)"},
    };

    worksheet.write_rich_string("A7", rich_string, center);
  }

  workbook.save("rich_strings.xlsx");
}
