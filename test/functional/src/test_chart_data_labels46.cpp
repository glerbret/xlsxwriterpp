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
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(74951296, 74965376);

  uint8_t data[5][4] = {
      {1, 2,  3,  10},
      {2, 4,  6,  20},
      {3, 6,  9,  30},
      {4, 8,  12, 40},
      {5, 10, 15, 50}
  };

  for(size_t row = 0; row < 5; row++)
  {
    for(size_t col = 0; col < 4; col++)
    {
      worksheet.write_number(row, col, data[row][col]);
    }
  }

  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$A$1:$A$5");

  xwpp::chart_line_t line = {.color_ = xwpp::color_t::RED};
  xwpp::chart_font_t font = {.color_ = xwpp::color_t::RED, .baseline_ = -1};
  xwpp::chart_fill_t fill = {.color_ = xwpp::color_t(0x00B050)};

  std::vector<xwpp::chart_data_label_t> data_labels = {
      {.value_ = "=Sheet1!$D$1", .font_ = font, .line_ = line, .fill_ = fill}
  };

  xwpp::chart_series_set_labels_custom(series, data_labels);

  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_data_labels46.xlsx");
}
