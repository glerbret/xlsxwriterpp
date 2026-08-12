/*
 * An example of creating an Excel pie chart using the Xlsxwriter++ library.
 *
 * The demo also shows how to set segment colors. It is possible to define
 * chart colors for most types of Xlsxwriter++ charts via the series
 * formatting functions. However, Pie/Doughnut charts are a special case since
 * each segment is represented as a point so it is necessary to assign
 * formatting to each point in the series.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

namespace
{

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  worksheet.write_string(CELL("A1"), "Category", bold);
  worksheet.write_string(CELL("A2"), "Apple");
  worksheet.write_string(CELL("A3"), "Cherry");
  worksheet.write_string(CELL("A4"), "Pecan");

  worksheet.write_string(CELL("B1"), "Values", bold);
  worksheet.write_number(CELL("B2"), 60);
  worksheet.write_number(CELL("B3"), 30);
  worksheet.write_number(CELL("B4"), 10);
}

}

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  write_worksheet_data(worksheet, bold);

  // Chart 1: Create a simple pie chart.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::PIE);

  // Add the first series to the chart.
  xwpp::chart_series_t& series1 = chart1.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart1.series_set_name(series1, "Pie sales data");

  // Add a chart title.
  chart1.title_set_name("Popular Pie Types");

  // Set an Excel chart style.
  chart1.set_style(10);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D2"), &chart1);

  // Chart 2: Create a pie chart with user defined segment colors.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::PIE);

  // Add the first series to the chart.
  xwpp::chart_series_t& series2 = chart2.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart2.series_set_name(series2, "Pie sales data");

  // Add a chart title.
  chart2.title_set_name("Pie Chart with user defined colors");

  // Add for fills for use in the chart.
  xwpp::chart_fill_t fill1{.color_ = static_cast<xwpp::color_t>(0x5ABA10)};
  xwpp::chart_fill_t fill2{.color_ = static_cast<xwpp::color_t>(0xFE110E)};
  xwpp::chart_fill_t fill3{.color_ = static_cast<xwpp::color_t>(0xCA5C05)};

  // Add some points with the above fills.
  const std::vector<xwpp::chart_point_t> points{
    {.fill_ = fill1},
    {.fill_ = fill2},
    {.fill_ = fill3},
  };

  // Add/override the points/segments of the chart.
  series_set_points(series2, points);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D18"), &chart2);

  // Chart 3: Create a pie chart with rotation of the segments.
  xwpp::chart_t& chart3 = workbook.add_chart(xwpp::chart_type_t::PIE);

  // Add the first series to the chart.
  xwpp::chart_series_t& series3 = chart3.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart3.series_set_name(series3, "Pie sales data");

  // Add a chart title.
  chart3.title_set_name("Pie Chart with segment rotation");

  // Change the angle/rotation of the first segment.
  chart3.set_rotation(90);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D34"), &chart3);

  workbook.save("chart_pie.xlsx");
}
