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
  const xwpp::format_t* merge_format = workbook.format_builder()
                                         .align(xwpp::format_horizontal_alignments_t::CENTER)
                                         .align(xwpp::format_vertical_alignments_t::CENTER)
                                         .border(xwpp::format_borders_t::THIN)
                                         .build();

  // Configure formats for the rich string.
  const xwpp::format_t* red  = workbook.format_builder().font_color(xwpp::color_t::red()).build();
  const xwpp::format_t* blue = workbook.format_builder().font_color(xwpp::color_t::blue()).build();

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
