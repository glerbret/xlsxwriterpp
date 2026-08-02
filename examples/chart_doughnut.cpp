/*
 * An example of creating an Excel doughnut chart using the Xlsxwriter++ library.
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

#include <vector>

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  worksheet.write_string(CELL("A1"), "Category", bold);
  worksheet.write_string(CELL("A2"), "Glazed");
  worksheet.write_string(CELL("A3"), "Chocolate");
  worksheet.write_string(CELL("A4"), "Cream");

  worksheet.write_string(CELL("B1"), "Values", bold);
  worksheet.write_number(CELL("B2"), 50);
  worksheet.write_number(CELL("B3"), 35);
  worksheet.write_number(CELL("B4"), 15);
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

  // Chart 1: Create a simple doughnut chart.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::DOUGHNUT);

  // Add the first series to the chart.
  xwpp::chart_series_t& series1 = chart1.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart1.series_set_name(series1, "Doughnut sales data");

  // Add a chart title.
  chart1.title_set_name("Popular Doughnut Types");

  // Set an Excel chart style.
  chart1.set_style(10);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D2"), &chart1);

  // Chart 2: Create a doughnut chart with user defined segment colors.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::DOUGHNUT);

  // Add the first series to the chart.
  xwpp::chart_series_t& series2 = chart2.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart2.series_set_name(series2, "Doughnut sales data");

  // Add a chart title.
  chart2.title_set_name("Doughnut Chart with user defined colors");

  // Add for fills for use in the chart.
  xwpp::chart_fill_t fill1{.color_ = static_cast<xwpp::color_t>(0xFA58D0)};
  xwpp::chart_fill_t fill2{.color_ = static_cast<xwpp::color_t>(0x61210B)};
  xwpp::chart_fill_t fill3{.color_ = static_cast<xwpp::color_t>(0xF5F6CE)};

  // Add some points with the above fills.
  xwpp::chart_point_t point1{.fill_ = fill1};
  xwpp::chart_point_t point2{.fill_ = fill2};
  xwpp::chart_point_t point3{.fill_ = fill3};

  // Create an array of the point objects.
  std::vector<xwpp::chart_point_t> points{
    point1,
    point2,
    point3,
  };

  // Add/override the points/segments of the chart.
  series_set_points(series2, points);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D18"), &chart2);

  // Chart 3: Create a Doughnut chart with rotation of the segments.
  xwpp::chart_t& chart3 = workbook.add_chart(xwpp::chart_type_t::DOUGHNUT);

  // Add the first series to the chart.
  xwpp::chart_series_t& series3 = chart3.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart3.series_set_name(series3, "Doughnut sales data");

  // Add a chart title.
  chart3.title_set_name("Doughnut Chart with segment rotation");

  // Change the angle/rotation of the first segment.
  chart3.set_rotation(90);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D34"), &chart3);

  // Chart 4: Create a Doughnut chart with user defined hole size and other options.
  xwpp::chart_t& chart4 = workbook.add_chart(xwpp::chart_type_t::DOUGHNUT);

  // Add the first series to the chart.
  xwpp::chart_series_t& series4 = chart4.add_series("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

  // Set the name for the series instead of the default "Series 1".
  chart4.series_set_name(series4, "Doughnut sales data");

  // Add a chart title.
  chart4.title_set_name("Doughnut Chart with options applied.");

  // Add/override the points/segments defined in Chart 2.
  series_set_points(series4, points);

  // Set an Excel chart style.
  chart4.set_style(26);

  // Change the angle/rotation of the first segment.
  chart4.set_rotation(28);

  // Change the hole size.
  chart4.set_hole_size(33);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("D50"), &chart4);

  workbook.save("chart_doughnut.xlsx");
}
