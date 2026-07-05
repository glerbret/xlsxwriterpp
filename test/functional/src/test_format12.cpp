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

  xwpp::format_t* top_left_bottom = workbook.add_format();
  top_left_bottom->set_bottom(xwpp::format_borders_t::THIN);
  top_left_bottom->set_left(xwpp::format_borders_t::THIN);
  top_left_bottom->set_top(xwpp::format_borders_t::THIN);

  xwpp::format_t* top_bottom = workbook.add_format();
  top_bottom->set_bottom(xwpp::format_borders_t::THIN);
  top_bottom->set_top(xwpp::format_borders_t::THIN);

  xwpp::format_t* top_left = workbook.add_format();
  top_left->set_left(xwpp::format_borders_t::THIN);
  top_left->set_top(xwpp::format_borders_t::THIN);

  xwpp::format_t* unused = workbook.add_format();
  unused->set_left(xwpp::format_borders_t::THIN);

  worksheet.write_string(CELL("B2"), "test", top_left_bottom);
  worksheet.write_string(CELL("D2"), "test", top_left);
  worksheet.write_string(CELL("F2"), "test", top_bottom);

  workbook.save("test_format12.xlsx");
}
