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

  worksheet.write_number(CELL("A1"), 10);
  worksheet.write_number(CELL("A2"), 20);
  worksheet.write_number(CELL("A3"), 30);
  worksheet.write_number(CELL("A4"), 40);

  xwpp::conditional_format_t conditional_format{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::THREE_ARROWS_COLORED,
  };

  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  workbook.save("test_cond_format17.xlsx");
}
