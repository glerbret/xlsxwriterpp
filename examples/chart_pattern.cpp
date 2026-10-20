/*
 * An example of a simple Excel chart with patterns using the Xlsxwriter++
 * library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format to use to highlight the header cells.
  const xwpp::format_t* bold = workbook.format_builder().bold().build();

  // Write some data for the chart.
  worksheet.write(0, 0, "Shingle", bold);
  worksheet.write(1, 0, 105);
  worksheet.write(2, 0, 150);
  worksheet.write(3, 0, 130);
  worksheet.write(4, 0, 90);
  worksheet.write(0, 1, "Brick", bold);
  worksheet.write(1, 1, 50);
  worksheet.write(2, 1, 120);
  worksheet.write(3, 1, 100);
  worksheet.write(4, 1, 110);

  // Create a chart object.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Configure the chart.
  xwpp::chart_series_t& series1 = chart.add_series("Sheet1!$A$2:$A$5");
  xwpp::chart_series_t& series2 = chart.add_series("Sheet1!$B$2:$B$5");

  series1.set_name("=Sheet1!$A$1");
  series2.set_name("=Sheet1!$B$1");

  chart.set_title("Cladding types");
  xwpp::chart_axis_set_name(chart.x_axis_, "Region");
  xwpp::chart_axis_set_name(chart.y_axis_, "Number of houses");

  // Configure an add the chart series patterns.
  const xwpp::chart_pattern_t pattern1{
    .fg_color = xwpp::color_t{0x804000},
    .bg_color = xwpp::color_t{0XC68C53},
    .type     = xwpp::chart_pattern_type_t::SHINGLE,
  };

  const xwpp::chart_pattern_t pattern2{
    .fg_color = xwpp::color_t{0XB30000},
    .bg_color = xwpp::color_t{0XFF6666},
    .type     = xwpp::chart_pattern_type_t::HORIZONTAL_BRICK,
  };

  series1.set_pattern(pattern1);
  series2.set_pattern(pattern2);

  // Configure and set the chart series borders.
  const xwpp::chart_line_t line2{.color = xwpp::color_t{0xb30000}};
  const xwpp::chart_line_t line1{.color = xwpp::color_t{0x804000}};

  series1.set_line(line1);
  series2.set_line(line2);

  // Widen the gap between the series/categories.
  chart.set_series_gap(70);

  // Insert the chart into the worksheet.
  worksheet.insert_chart("D2", &chart);

  workbook.save("chart_pattern.xlsx");
}
