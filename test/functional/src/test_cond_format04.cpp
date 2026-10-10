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

  const xwpp::format_t* format1 = workbook.format_builder().num_format_index(2).build();
  const xwpp::format_t* format2 = workbook.format_builder().num_format("0.000").build();

  // We manually set the indices to get the same order as the target file.
  (void)format2->get_dxf_index();
  (void)format1->get_dxf_index();

  worksheet.write("A1", 10);
  worksheet.write("A2", 20);
  worksheet.write("A3", 30);
  worksheet.write("A4", 40);

  const xwpp::conditional_format_t conditional_format1{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN,
    .value_    = 2,
    .format_   = format1,
  };
  worksheet.conditional_format_cell("A1", conditional_format1);

  const xwpp::conditional_format_t conditional_format2{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
    .value_    = 8,
    .format_   = format2,
  };
  worksheet.conditional_format_cell("A2", conditional_format2);

  workbook.save("test_cond_format04.xlsx");
}
