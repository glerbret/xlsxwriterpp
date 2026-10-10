/*
 * An example of merging cells using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  // Configure a format for the merged range.
  const xwpp::format_t* merge_format =
    workbook.format_builder()
      .align(xwpp::format_horizontal_alignments_t::CENTER, xwpp::format_vertical_alignments_t::CENTER)
      .bold()
      .bg_color(xwpp::color_t::yellow())
      .border(xwpp::format_borders_t::THIN)
      .build();

  // Increase the cell size of the merged cells to highlight the formatting.
  worksheet.set_column(1, 3, 12);
  worksheet.set_row(3, 30);
  worksheet.set_row(6, 30);
  worksheet.set_row(7, 30);

  // Merge 3 cells.
  worksheet.merge_range(3, 1, 3, 3, "Merged Range", merge_format);

  // Merge 3 cells over two rows.
  worksheet.merge_range(6, 1, 7, 3, "Merged Range", merge_format);

  workbook.save("merge_range.xlsx");
}
