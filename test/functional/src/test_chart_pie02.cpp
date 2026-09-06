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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

  const std::array<std::array<uint8_t, 2>, 3> data{
    {{2, 60}, {4, 30}, {6, 10}}
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  chart.add_series("=Sheet1!$A$1:$A$3", "=Sheet1!$B$1:$B$3");

  const xwpp::chart_font_t font{.bold_ = true, .italic_ = true, .baseline_ = -1};
  chart.legend_set_font(font);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_pie02.xlsx");
}
