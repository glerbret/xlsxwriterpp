/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bg_color(xwpp::color_t(0xFF0000));

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_bg_color(xwpp::color_t(0x92D050));

  const std::array<std::array<uint8_t, 10>, 10> data{
    {
     {90, 80, 50, 10, 20, 90, 40, 90, 30, 40},
     {20, 10, 90, 100, 30, 60, 70, 60, 50, 90},
     {10, 50, 60, 50, 20, 50, 80, 30, 40, 60},
     {10, 90, 20, 40, 10, 40, 50, 70, 90, 50},
     {70, 100, 10, 90, 10, 10, 20, 100, 100, 40},
     {20, 60, 10, 100, 30, 10, 20, 60, 100, 10},
     {10, 60, 10, 80, 100, 80, 30, 30, 70, 40},
     {30, 90, 60, 10, 10, 100, 40, 40, 30, 40},
     {80, 90, 10, 20, 20, 50, 80, 20, 60, 90},
     {60, 80, 30, 30, 10, 50, 80, 60, 50, 30},
     }
  };

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  // We manually set the indices to get the same order as the target file.
  format2->get_dxf_index_(format2);
  format1->get_dxf_index_(format1);

  const xwpp::conditional_format_t conditional_format1{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO,
    .value_    = 50,
    .format_   = format1,
  };
  worksheet.conditional_format_range(RANGE("A1:J10"), conditional_format1);

  const xwpp::conditional_format_t conditional_format2{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
    .value_    = 50,
    .format_   = format2,
  };
  worksheet.conditional_format_range(RANGE("A1:J10"), conditional_format2);

  workbook.save("test_cond_format07.xlsx");
}
