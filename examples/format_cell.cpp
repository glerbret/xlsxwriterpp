/*
 * Example of writing some data with cell formatting to a simple Excel
 * file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Widen the first column to make the text clearer.
  worksheet.set_column(1, 1, 30);

  {
    const xwpp::format_t* format = workbook.format_builder().bg_color(xwpp::color_t::yellow()).build();

    worksheet.write(1, 1, "Yellow cell", format);
  }

  {
    const xwpp::format_t* format =
      workbook.format_builder().border(xwpp::format_borders_t::MEDIUM).border_color(xwpp::color_t::red()).build();

    worksheet.write(3, 1, "Cell with red borders", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder()
                                     .bottom(xwpp::format_borders_t::DASHED)
                                     .bottom_color(xwpp::color_t::yellow())
                                     .top(xwpp::format_borders_t::DOTTED)
                                     .top_color(xwpp::color_t::red())
                                     .left(xwpp::format_borders_t::THICK)
                                     .left_color(xwpp::color_t::blue())
                                     .right(xwpp::format_borders_t::DOUBLE)
                                     .right_color(xwpp::color_t::green())
                                     .build();

    worksheet.write(5, 1, "Cell with different borders", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder()
                                     .diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN)
                                     .diag_border(xwpp::format_borders_t::THICK)
                                     .diag_color(xwpp::color_t::blue())
                                     .build();

    worksheet.write(7, 1, "Cell with diag", format);
  }

  workbook.save("format_cell.xlsx");
}
