/*
 * An example of merging cells containing a rich string using libxlsxwriter.
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

  // Configure a format for the merged range.
  xwpp::format_t* merge_format = workbook.add_format();
  merge_format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  merge_format->set_align(xwpp::format_alignments_t::VERTICAL_CENTER);
  merge_format->set_border(xwpp::format_borders_t::THIN);

  // Configure formats for the rich string.
  xwpp::format_t* red = workbook.add_format();
  red->set_font_color(xwpp::color_t::RED);

  xwpp::format_t* blue = workbook.add_format();
  blue->set_font_color(xwpp::color_t::BLUE);

  // Create the fragments for the rich string.

  const std::vector<xwpp::rich_string_tuple_t> rich_string{
    {.str_ = "This is "},
    {.format_ = red, .str_ = "red"},
    {.str_ = " and this is "},
    {.format_ = blue, .str_ = "blue"},
  };

  // Write an empty string to the merged range.
  worksheet.merge_range(1, 1, 4, 3, "", merge_format);

  // We then overwrite the first merged cell with a rich string. Note that
  // we must also pass the cell format used in the merged cells format at
  // the end.
  worksheet.write_rich_string(1, 1, rich_string, merge_format);

  workbook.save("merge_rich_string.xlsx");
}
