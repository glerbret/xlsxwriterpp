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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45861120, 45867008);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
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

  chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  const xwpp::chart_layout_t layout{
    .x_ = 0.359652668416448,
    .y_ = 0.1388888888888889,
  };

  chart.title_set_name_range("Sheet1", 0, 0);
  chart.title_set_layout(layout);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_layout07.xlsx");
}
