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
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Panes 1");

    worksheet.freeze_panes(1, 0);

    // Some sheet formatting.
    worksheet.set_column(0, 8, 16);
    worksheet.set_row(0, 20);
    worksheet.set_selection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    for(xwpp::col_num_t col{0}; col < 9; col++)
    {
      worksheet.write_string(0, col, "Scroll down", header);
    }

    for(xwpp::row_num_t row_num{1}; row_num < 100; row_num++)
    {
      for(xwpp::col_num_t col_num{0}; col_num < 9; col_num++)
      {
        worksheet.write_number(row_num, col_num, row_num + 1, center);
      }
    }
  }

  // Example 2. Freeze pane on the left column.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Panes 2");

    worksheet.freeze_panes(0, 1);

    // Some sheet formatting.
    worksheet.set_column(0, 0, 16);
    worksheet.set_selection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    for(xwpp::row_num_t row_num{0}; row_num < 50; row_num++)
    {
      worksheet.write_string(row_num, 0, "Scroll right", header);

      for(xwpp::col_num_t col_num{1}; col_num < 26; col_num++)
      {
        worksheet.write_number(row_num, col_num, col_num, center);
      }
    }
  }

  // Example 3. Freeze pane on the top row and left column.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Panes 3");

    worksheet.freeze_panes(1, 1);

    // Some sheet formatting.
    worksheet.set_column(0, 25, 16);
    worksheet.set_row(0, 20);
    worksheet.write_string(0, 0, "", header);
    worksheet.set_selection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    for(xwpp::col_num_t col{1}; col < 26; col++)
    {
      worksheet.write_string(0, col, "Scroll down", header);
    }

    for(xwpp::row_num_t row_num{1}; row_num < 50; row_num++)
    {
      worksheet.write_string(row_num, 0, "Scroll right", header);

      for(xwpp::col_num_t col_num{1}; col_num < 26; col_num++)
      {
        worksheet.write_number(row_num, col_num, col_num, center);
      }
    }
  }

  // Example 4. Split pane on the top row and left column.
  //
  // The divisions must be specified in terms of row and column dimensions.
  // The default row height is 15 and the default column width is 8.43
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Panes 4");

    worksheet.split_panes(15, 8.43);

    // Some worksheet text to demonstrate scrolling.
    for(xwpp::col_num_t col{1}; col < 26; col++)
    {
      worksheet.write_string(0, col, "Scroll", center);
    }

    for(xwpp::row_num_t row_num{1}; row_num < 50; row_num++)
    {
      worksheet.write_string(row_num, 0, "Scroll", center);

      for(xwpp::col_num_t col_num{1}; col_num < 26; col_num++)
      {
        worksheet.write_number(row_num, col_num, col_num, center);
      }
    }
  }

  workbook.save("panes.xlsx");
}
