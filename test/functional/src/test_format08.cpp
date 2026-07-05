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
  xwpp::format_t* border5 = workbook.add_format();

  border1->set_bottom(xwpp::format_borders_t::THIN);
  border1->set_bottom_color(xwpp::color_t::RED);

  border2->set_top(xwpp::format_borders_t::THIN);
  border2->set_top_color(xwpp::color_t::RED);

  border3->set_left(xwpp::format_borders_t::THIN);
  border3->set_left_color(xwpp::color_t::RED);

  border4->set_right(xwpp::format_borders_t::THIN);
  border4->set_right_color(xwpp::color_t::RED);

  border5->set_border(xwpp::format_borders_t::THIN);
  border5->set_border_color(xwpp::color_t::RED);

  worksheet.write_blank(1, 1, border1);
  worksheet.write_blank(3, 1, border2);
  worksheet.write_blank(5, 1, border3);
  worksheet.write_blank(7, 1, border4);
  worksheet.write_blank(9, 1, border5);

  workbook.save("test_format08.xlsx");
}
