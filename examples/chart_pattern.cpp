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
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  worksheet.write_string(0, 0, "Shingle", bold);
  worksheet.write_number(1, 0, 105);
  worksheet.write_number(2, 0, 150);
  worksheet.write_number(3, 0, 130);
  worksheet.write_number(4, 0, 90);
  worksheet.write_string(0, 1, "Brick", bold);
  worksheet.write_number(1, 1, 50);
  worksheet.write_number(2, 1, 120);
  worksheet.write_number(3, 1, 100);
  worksheet.write_number(4, 1, 110);

  // Create a chart object.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Configure the chart.
  xwpp::chart_series_t& series1 = chart.add_series("", "Sheet1!$A$2:$A$5");
  xwpp::chart_series_t& series2 = chart.add_series("", "Sheet1!$B$2:$B$5");

  chart.series_set_name(series1, "=Sheet1!$A$1");
  chart.series_set_name(series2, "=Sheet1!$B$1");

  chart.title_set_name("Cladding types");
  xwpp::chart_axis_set_name(chart.x_axis_, "Region");
  xwpp::chart_axis_set_name(chart.y_axis_, "Number of houses");

  // Configure an add the chart series patterns.
  xwpp::chart_pattern_t pattern1{
    .fg_color_ = static_cast<xwpp::color_t>(0x804000),
    .bg_color_ = static_cast<xwpp::color_t>(0XC68C53),
    .type_     = xwpp::chart_pattern_type_t::SHINGLE,
  };

  xwpp::chart_pattern_t pattern2{
    .fg_color_ = static_cast<xwpp::color_t>(0XB30000),
    .bg_color_ = static_cast<xwpp::color_t>(0XFF6666),
    .type_     = xwpp::chart_pattern_type_t::HORIZONTAL_BRICK,
  };

  xwpp::chart_series_set_pattern(series1, pattern1);
  xwpp::chart_series_set_pattern(series2, pattern2);

  // Configure and set the chart series borders.
  xwpp::chart_line_t line1{.color_ = static_cast<xwpp::color_t>(0x804000)};
  xwpp::chart_line_t line2{.color_ = static_cast<xwpp::color_t>(0xb30000)};

  xwpp::chart_series_set_line(series1, line1);
  xwpp::chart_series_set_line(series2, line2);

  // Widen the gap between the series/categories.
  chart.set_series_gap(70);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D2"), &chart);

  workbook.save("chart_pattern.xlsx");
}
