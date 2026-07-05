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

  uint8_t data[5][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {5, 10, 15}
  };

  std::vector<std::string> sheetnames{
      "Sheet 1", "Sheet 2", "Sheet!3", "Sheet\"4", "Sheet#5", "Sheet$6", "Sheet%7", "Sheet&8",
  };

  for(const auto& name: sheetnames)
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet(name);

    for(size_t row = 0; row < 5; row++)
    {
      for(size_t col = 0; col < 3; col++)
      {
        worksheet.write_number(row, col, data[row][col]);
      }
    }

    xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);
    xwpp::chart_series_t& series = chart.add_series("", "");
    chart_series_set_values(series, name, 0, 0, 4, 0);

    xwpp::chart_options_t options = {.x_offset_ = 26, .y_offset_ = 17};
    worksheet.insert_chart(CELL("E6"), &chart, options);
  }

  workbook.save("test_quote_name01.xlsx");
}
