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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

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

  xwpp::chart_series_t& series1 = chart.add_series("", "=Sheet1!$A$1:$A$5");

  xwpp::chart_font_t font = {.name_ = "Consolas", .pitch_family_ = 49, .charset_ = 0, .baseline_ = -1};
  xwpp::chart_line_t line = {.color_ = xwpp::color_t::RED};

  xwpp::chart_series_set_labels(series1);
  xwpp::chart_series_set_labels_options(series1, true, true, true);
  xwpp::chart_series_set_labels_position(series1, xwpp::chart_label_position_t::INSIDE_END);
  xwpp::chart_series_set_labels_separator(series1, xwpp::chart_label_separator_t::SEMICOLON);
  xwpp::chart_series_set_labels_leader_line(series1);
  xwpp::chart_series_set_labels_legend(series1);
  xwpp::chart_series_set_labels_percentage(series1);
  xwpp::chart_series_set_labels_num_format(series1, "#,##0.00");
  xwpp::chart_series_set_labels_font(series1, font);
  xwpp::chart_series_set_labels_line(series1, line);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_data_labels36.xlsx");
}
