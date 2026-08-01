/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(68809856, 68811392);

  uint8_t data[5][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {3, 6,  9 },
      {4, 8,  12},
      {5, 10, 15}
  };

  for(xwpp::row_num_t row = 0; row < 5; row++)
  {
    for(xwpp::col_num_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series1 =
      chart.add_series("=(Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)", "=(Sheet1!$B$1:$B$2,Sheet1!$B$4:$B$5)");

  worksheet.insert_chart(CELL("E9"), &chart);

  /* Add the cached data for testing. */
  uint8_t test_data[4][3] = {
      {1, 2,  3 },
      {2, 4,  6 },
      {4, 8,  12},
      {5, 10, 15}
  };

  xwpp::chart_add_data_cache(series1.categories_, test_data[0], 4, 3, 0);
  xwpp::chart_add_data_cache(series1.values_, test_data[0], 4, 3, 1);

  workbook.save("test_chart_column08.xlsx");
}
