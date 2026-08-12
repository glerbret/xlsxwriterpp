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

  const std::vector<std::string> sheetnames{
    "Sheet 1", "Sheet 2", "Sheet!3", "Sheet\"4", "Sheet#5", "Sheet$6", "Sheet%7", "Sheet&8",
  };

  for(const auto& name: sheetnames)
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet(name);

    for(xwpp::row_num_t row_num = 0; const auto& row: data)
    {
      for(xwpp::col_num_t col_num = 0; const auto value: row)
      {
        worksheet.write_number(row_num, col_num, value);
        col_num++;
      }
      row_num++;
    }

    xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);
    xwpp::chart_series_t& series = chart.add_series("", "");
    chart_series_set_values(series, name, 0, 0, 4, 0);

    const xwpp::chart_options_t options{.x_offset_ = 26, .y_offset_ = 17};
    worksheet.insert_chart(CELL("E6"), &chart, options);
  }

  workbook.save("test_quote_name01.xlsx");
}
