/*
 * Example of writing some data with cell formatting to a simple Excel
 * file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Widen the first column to make the text clearer.
  worksheet.set_column(1, 1, 30);

  // Add some formats.
  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bg_color(xwpp::color_t::YELLOW);

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_border(xwpp::format_borders_t::MEDIUM);
  format2->set_border_color(xwpp::color_t::RED);

  xwpp::format_t* format3 = workbook.add_format();
  format3->set_bottom(xwpp::format_borders_t::DASHED);
  format3->set_bottom_color(xwpp::color_t::YELLOW);
  format3->set_top(xwpp::format_borders_t::DOTTED);
  format3->set_top_color(xwpp::color_t::RED);
  format3->set_left(xwpp::format_borders_t::THICK);
  format3->set_left_color(xwpp::color_t::BLUE);
  format3->set_right(xwpp::format_borders_t::DOUBLE);
  format3->set_right_color(xwpp::color_t::GREEN);

  xwpp::format_t* format4 = workbook.add_format();
  format4->set_diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN);
  format4->set_diag_border(xwpp::format_borders_t::THICK);
  format4->set_diag_color(xwpp::color_t::BLUE);

  worksheet.write_string(1, 1, "Yellow cell", format1);
  worksheet.write_string(3, 1, "Cell with red borders", format2);
  worksheet.write_string(5, 1, "Cell with different borders", format3);
  worksheet.write_string(7, 1, "Cell with diag", format4);

  workbook.save("format_cell.xlsx");
}
