/*
 * A simple formatting example that demonstrates how to add diagonal
 * cell borders using the Xlsxwriter++ library.
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

  // Add some diagonal border formats.
  xwpp::format_t* format1 = workbook.add_format();
  format1->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP);

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN);

  xwpp::format_t* format3 = workbook.add_format();
  format3->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN);

  xwpp::format_t* format4 = workbook.add_format();
  format4->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN);
  format4->set_diag_border(xwpp::format_borders_t::HAIR);
  format4->set_diag_color(xwpp::color_t::RED);

  worksheet.write_string(CELL("B3"),  "Text", format1);
  worksheet.write_string(CELL("B6"),  "Text", format2);
  worksheet.write_string(CELL("B9"),  "Text", format3);
  worksheet.write_string(CELL("B12"), "Text", format4);

  workbook.save("diagonal_border.xlsx");
}
