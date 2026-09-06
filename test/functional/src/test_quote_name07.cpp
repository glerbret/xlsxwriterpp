/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>
#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;

  const std::string sheetname{"Sheet'1"};
  xwpp::worksheet_t& worksheet = workbook.add_worksheet(sheetname);
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(48135552, 54701056);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
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

  worksheet.repeat_rows(0, 1);
  worksheet.set_portrait();
  worksheet.set_dpi(0, 200);

  xwpp::chart_series_t& series1 = chart.add_series("", "");
  xwpp::chart_series_t& series2 = chart.add_series("", "");
  xwpp::chart_series_t& series3 = chart.add_series("", "");

  chart_series_set_values(series1, sheetname, 0, 0, 4, 0);
  chart_series_set_values(series2, sheetname, 0, 1, 4, 1);
  chart_series_set_values(series3, sheetname, 0, 2, 4, 2);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_quote_name07.xlsx");
}
