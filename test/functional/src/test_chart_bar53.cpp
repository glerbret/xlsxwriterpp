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
  xwpp::chart_t& chart1        = workbook.add_chart(xwpp::chart_type_t::BAR);
  xwpp::chart_t& chart2        = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart1.set_axis_ids(64265216, 64447616);
  chart2.set_axis_ids(86048128, 86058112);

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

  xwpp::chart_series_t& series1 = chart1.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");
  xwpp::chart_series_t& series2 = chart1.add_series("=Sheet1!$A$1:$A$5", "=Sheet1!$C$1:$C$5");
  worksheet.insert_chart(CELL("E9"), &chart1);

  xwpp::chart_series_t& series3 = chart2.add_series("=Sheet1!$A$1:$A$4", "=Sheet1!$B$1:$B$4");
  xwpp::chart_series_t& series4 = chart2.add_series("=Sheet1!$A$1:$A$4", "=Sheet1!$C$1:$C$4");
  worksheet.insert_chart(CELL("F25"), &chart2);

  // Add cache data for testing.
  xwpp::chart_add_data_cache(series1.categories_, data[0].data(), 5, 3, 0);
  xwpp::chart_add_data_cache(series2.categories_, data[0].data(), 5, 3, 0);
  xwpp::chart_add_data_cache(series1.values_, data[0].data(), 5, 3, 1);
  xwpp::chart_add_data_cache(series2.values_, data[0].data(), 5, 3, 2);
  xwpp::chart_add_data_cache(series3.categories_, data[0].data(), 4, 3, 0);
  xwpp::chart_add_data_cache(series4.categories_, data[0].data(), 4, 3, 0);
  xwpp::chart_add_data_cache(series3.values_, data[0].data(), 4, 3, 1);
  xwpp::chart_add_data_cache(series4.values_, data[0].data(), 4, 3, 2);

  workbook.save("test_chart_bar53.xlsx");
}
