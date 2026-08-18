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

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  const std::string sheetname  = "Sheet 1";
  xwpp::worksheet_t& worksheet = workbook.add_worksheet(sheetname);

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::PIE);

  xwpp::chart_series_t& series = chart.add_series("", "");
  chart_series_set_values(series, sheetname, 0, 0, 4, 0);
  chart_series_set_name_range(series, sheetname, 0, 0);

  chart.title_set_name("Foo");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_quote_name04.xlsx");
}