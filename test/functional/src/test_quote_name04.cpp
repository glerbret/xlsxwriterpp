/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

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

  std::string sheetname        = "Sheet 1";
  xwpp::worksheet_t& worksheet = workbook.add_worksheet(sheetname);

  for(xwpp::row_num_t row = 0; row < 5; row++)
  {
    for(xwpp::col_num_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::PIE);

  xwpp::chart_series_t& series = chart.add_series("", "");
  chart_series_set_values(series, sheetname, 0, 0, 4, 0);
  chart_series_set_name_range(series, sheetname, 0, 0);

  chart.title_set_name("Foo");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_quote_name04.xlsx");
}