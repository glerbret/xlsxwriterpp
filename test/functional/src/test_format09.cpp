/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* border1 = workbook.add_format();
  xwpp::format_t* border2 = workbook.add_format();
  xwpp::format_t* border3 = workbook.add_format();
  xwpp::format_t* border4 = workbook.add_format();

  border1->set_border(xwpp::format_borders_t::HAIR);
  border1->set_border_color(xwpp::color_t::RED);

  border2->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP);
  border2->set_diag_color(xwpp::color_t::RED);

  border3->set_diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN);
  border3->set_diag_color(xwpp::color_t::RED);

  border4->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN);
  border4->set_diag_color(xwpp::color_t::RED);

  worksheet.write_blank(1, 1, border1);
  worksheet.write_blank(3, 1, border2);
  worksheet.write_blank(5, 1, border3);
  worksheet.write_blank(7, 1, border4);

  workbook.save("test_format09.xlsx");
}
