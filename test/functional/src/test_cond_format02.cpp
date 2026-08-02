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

  worksheet.write_number(CELL("A1"), 10, NULL);
  worksheet.write_number(CELL("A2"), 20, NULL);
  worksheet.write_number(CELL("A3"), 30, NULL);
  worksheet.write_number(CELL("A4"), 40, NULL);

  xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
    .value_    = 5,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  workbook.save("test_cond_format02.xlsx");
}
