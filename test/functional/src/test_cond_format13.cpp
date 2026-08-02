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

  xwpp::format_t* format1 = workbook.add_format();
  format1->set_num_format_index(2);

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_num_format("0.000");

  // We manually set the indices to get the same order as the target file.
  format2->get_dxf_index_(format2);
  format1->get_dxf_index_(format1);

  worksheet.write_number(CELL("A1"), 10, NULL);
  worksheet.write_number(CELL("A2"), 20, NULL);
  worksheet.write_number(CELL("A3"), 30, NULL);
  worksheet.write_number(CELL("A4"), 40, NULL);

  xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN,
    .value_    = 2,
    .format_   = format1,
  };
  worksheet.conditional_format_cell(CELL("A1"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::LESS_THAN;
  conditional_format.value_    = 8;
  conditional_format.format_   = format2;
  worksheet.conditional_format_cell(CELL("A2"), conditional_format);

  workbook.save("test_cond_format13.xlsx");
}
