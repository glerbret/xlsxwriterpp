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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(45740032, 45743104);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  const xwpp::chart_font_t font{.name = "Consolas", .pitch_family = 49, .charset = 0, .baseline = -1};

  chart.add_series("=Sheet1!$A$1:$A$5")
    .set_labels()
    .set_labels_options(true, true, true)
    .set_labels_position(xwpp::chart_label_position_t::INSIDE_END)
    .set_labels_separator(xwpp::chart_label_separator_t::SEMICOLON)
    .set_labels_leader_line()
    .set_labels_legend()
    .set_labels_percentage()
    .set_labels_num_format("#,##0.00")
    .set_labels_font(font);

  worksheet.insert_chart("E9", &chart);

  workbook.save("test_chart_data_labels21.xlsx");
}
