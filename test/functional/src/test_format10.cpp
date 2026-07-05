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

  border1->set_bg_color(xwpp::color_t::RED);

  border2->set_bg_color(xwpp::color_t::YELLOW);
  border2->set_pattern(xwpp::format_patterns_t::DARK_VERTICAL);

  border3->set_bg_color(xwpp::color_t::YELLOW);
  border3->set_fg_color(xwpp::color_t::RED);
  border3->set_pattern(xwpp::format_patterns_t::GRAY_0625);

  worksheet.write_blank(1, 1, border1);
  worksheet.write_blank(3, 1, border2);
  worksheet.write_blank(5, 1, border3);

  workbook.save("test_format10.xlsx");
}
