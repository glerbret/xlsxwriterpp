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

  xwpp::format_t* format = workbook.add_format();
  format->set_pattern(xwpp::format_patterns_t::LIGHT_VERTICAL);
  format->set_bg_color(xwpp::color_t(0xFFFF00));
  format->set_fg_color(xwpp::color_t(0xFF0000));

  worksheet.write_string(CELL("A1"), "Hello", format);

  worksheet.write_number(CELL("B3"), 10);
  worksheet.write_number(CELL("B4"), 20);
  worksheet.write_number(CELL("B5"), 30);
  worksheet.write_number(CELL("B6"), 40);

  const xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN,
    .value_    = 20,
    .format_   = format,
  };
  worksheet.conditional_format_range(RANGE("B3:B6"), conditional_format);

  workbook.save("test_cond_format12.xlsx");
}
