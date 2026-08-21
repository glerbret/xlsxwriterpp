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
    xwpp::format_t* format = workbook.add_format();
    format->set_bg_color(xwpp::color_t::YELLOW);

    worksheet.write_string(1, 1, "Yellow cell", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_border(xwpp::format_borders_t::MEDIUM);
    format->set_border_color(xwpp::color_t::RED);

    worksheet.write_string(3, 1, "Cell with red borders", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_bottom(xwpp::format_borders_t::DASHED);
    format->set_bottom_color(xwpp::color_t::YELLOW);
    format->set_top(xwpp::format_borders_t::DOTTED);
    format->set_top_color(xwpp::color_t::RED);
    format->set_left(xwpp::format_borders_t::THICK);
    format->set_left_color(xwpp::color_t::BLUE);
    format->set_right(xwpp::format_borders_t::DOUBLE);
    format->set_right_color(xwpp::color_t::GREEN);

    worksheet.write_string(5, 1, "Cell with different borders", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN);
    format->set_diag_border(xwpp::format_borders_t::THICK);
    format->set_diag_color(xwpp::color_t::BLUE);

    worksheet.write_string(7, 1, "Cell with diag", format);
  }

  workbook.save("format_cell.xlsx");
}
