/*
 * A simple example using the Xlsxwriter++ library to create worksheets with
 * panes.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add some worksheets.
  xwpp::workbook_t workbook;

  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet("Panes 1");
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet("Panes 2");
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet("Panes 3");
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet("Panes 4");

  // Set up some formatting and text to highlight the panes.
  xwpp::format_t* header = workbook.add_format();
  header->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  header->set_align(xwpp::format_alignments_t::VERTICAL_CENTER);
  header->set_fg_color(xwpp::color_t(0xD7E4BC));
  header->set_bold();
  header->set_border(xwpp::format_borders_t::THIN);

  xwpp::format_t* center = workbook.add_format();
  center->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);

  // Example 1. Freeze pane on the top row.
  worksheet1.freeze_panes(1, 0);

  // Some sheet formatting.
  worksheet1.set_column(0, 8, 16);
  worksheet1.set_row(0, 20);
  worksheet1.set_selection(4, 3, 4, 3);

  // Some worksheet text to demonstrate scrolling.
  for(xwpp::col_num_t col = 0; col < 9; col++)
  {
    worksheet1.write_string(0, col, "Scroll down", header);
  }

  for(xwpp::row_num_t row = 1; row < 100; row++)
  {
    for(xwpp::col_num_t col = 0; col < 9; col++)
    {
      worksheet1.write_number(row, col, row + 1, center);
    }
  }

  // Example 2. Freeze pane on the left column.
  worksheet2.freeze_panes(0, 1);

  // Some sheet formatting.
  worksheet2.set_column(0, 0, 16);
  worksheet2.set_selection(4, 3, 4, 3);

  // Some worksheet text to demonstrate scrolling.
  for(xwpp::row_num_t row = 0; row < 50; row++)
  {
    worksheet2.write_string(row, 0, "Scroll right", header);

    for(xwpp::col_num_t col = 1; col < 26; col++)
    {
      worksheet2.write_number(row, col, col, center);
    }
  }

  // Example 3. Freeze pane on the top row and left column.
  worksheet3.freeze_panes(1, 1);

  // Some sheet formatting.
  worksheet3.set_column(0, 25, 16);
  worksheet3.set_row(0, 20);
  worksheet3.write_string(0, 0, "", header);
  worksheet3.set_selection(4, 3, 4, 3);

  // Some worksheet text to demonstrate scrolling.
  for(xwpp::col_num_t col = 1; col < 26; col++)
  {
    worksheet3.write_string(0, col, "Scroll down", header);
  }

  for(xwpp::row_num_t row = 1; row < 50; row++)
  {
    worksheet3.write_string(row, 0, "Scroll right", header);

    for(xwpp::col_num_t col = 1; col < 26; col++)
    {
      worksheet3.write_number(row, col, col, center);
    }
  }

  /* Example 4. Split pane on the top row and left column.
   *
   * The divisions must be specified in terms of row and column dimensions.
   * The default row height is 15 and the default column width is 8.43
   */
  worksheet4.split_panes(15, 8.43);

  // Some sheet formatting.

  // Some worksheet text to demonstrate scrolling.
  for(xwpp::col_num_t col = 1; col < 26; col++)
  {
    worksheet4.write_string(0, col, "Scroll", center);
  }

  for(xwpp::row_num_t row = 1; row < 50; row++)
  {
    worksheet4.write_string(row, 0, "Scroll", center);

    for(xwpp::col_num_t col = 1; col < 26; col++)
    {
      worksheet4.write_number(row, col, col, center);
    }
  }

  workbook.save("panes.xlsx");
}
