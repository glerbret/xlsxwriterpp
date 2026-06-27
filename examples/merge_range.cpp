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
  xwpp::format_t* merge_format = workbook.add_format();

  // Configure a format for the merged range.
  merge_format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  merge_format->set_align(xwpp::format_alignments_t::VERTICAL_CENTER);
  merge_format->set_bold();
  merge_format->set_bg_color(xwpp::color_t::YELLOW);
  merge_format->set_border(xwpp::format_borders_t::THIN);

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
