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
  format->set_font_color(xwpp::color_t(0x9C6500));
  format->set_bg_color(xwpp::color_t(0xFFEB9C));
  format->set_font_condense();
  format->set_font_extend();

  worksheet.write_number(CELL("A1"), 10);
  worksheet.write_number(CELL("A2"), 20);
  worksheet.write_number(CELL("A3"), 30);
  worksheet.write_number(CELL("A4"), 40);

  const xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN,
    .value_    = 5,
    .format_   = format,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  workbook.save("test_cond_format09.xlsx");
}
