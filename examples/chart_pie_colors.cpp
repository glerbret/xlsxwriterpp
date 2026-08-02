/*
 * An example of creating an Excel pie chart with user defined colors using
 * the Xlsxwriter++ library.
 *
 * In general formatting is applied to an entire series in a chart. However,
 * it is occasionally required to format individual points in a series. In
 * particular this is required for Pie/Doughnut charts where each segment is
 * represented by a point.
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

  // Write some data for the chart.
  worksheet.write_string(CELL("A1"), "Pass");
  worksheet.write_string(CELL("A2"), "Fail");
  worksheet.write_number(CELL("B1"), 90);
  worksheet.write_number(CELL("B2"), 10);

  // Create a pie chart.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::PIE);

  // Add the data series to the chart.
  xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$1:$A$2", "=Sheet1!$B$1:$B$2");

  // Create some fills for the chart points/segments.
  xwpp::chart_fill_t red_fill{.color_ = xwpp::color_t::RED};
  xwpp::chart_fill_t green_fill{.color_ = xwpp::color_t::GREEN};

  // Add the fills to the point objects.
  std::vector<xwpp::chart_point_t> points{
    {.fill_ = green_fill},
    {.fill_ = red_fill},
  };

  // Add the points to the series. */
  series_set_points(series, points);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D2"), &chart);

  workbook.save("chart_pie_colors.xlsx");
}
