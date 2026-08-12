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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::BAR);

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

  xwpp::chart_series_t series1 = chart.add_series("", "Sheet1!$A$1:$A$5");
  xwpp::chart_series_t series2 = chart.add_series("", "Sheet1!$B$1:$B$5");

  // Add the cached data for testing.
  xwpp::chart_add_data_cache(series1.values_, data[0].data(), 5, 3, 0);
  xwpp::chart_add_data_cache(series2.values_, data[0].data(), 5, 3, 1);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_bar51.xlsx");
}
