/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format = workbook.add_format();

  format->set_pattern(xwpp::format_patterns_t::LIGHT_GRID);
  format->set_bg_color(xwpp::color_t(0xFFFF00));
  format->set_fg_color(xwpp::color_t(0xFF0000));

  worksheet.write_number(CELL("A1"), 10, NULL);
  worksheet.write_number(CELL("A2"), 20, NULL);
  worksheet.write_number(CELL("A3"), 30, NULL);
  worksheet.write_number(CELL("A4"), 40, NULL);

  xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN,
    .value_    = 7,
    .format_   = format,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  workbook.save("test_cond_format06.xlsx");
}
