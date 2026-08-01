/*
 * An example of creating an Excel line chart using the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  uint8_t data[6][3] = {
      {2, 10, 30},
      {3, 40, 60},
      {4, 50, 70},
      {5, 20, 50},
      {6, 10, 40},
      {7, 50, 30}
  };

  worksheet.write_string(CELL("A1"), "Number", bold);
  worksheet.write_string(CELL("B1"), "Batch 1", bold);
  worksheet.write_string(CELL("C1"), "Batch 2", bold);

  for(xwpp::row_num_t row = 0; row < 6; row++)
  {
    for(xwpp::col_num_t col = 0; col < 3; col++)
    {
      worksheet.write_number(row + 1, col, data[row][col]);
    }
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

  // Chart 1. Create a area chart.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add the first series to the chart.
  xwpp::chart_series_t& series1 = chart1.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart1.series_set_name(series1, "=Sheet1!$B$1");

  // Add a second series but leave the categories and values undefined. They
  // can be defined later using the alternative syntax shown below.  */
  xwpp::chart_series_t& series2 = chart1.add_series("", "");

  // Configure the series using a syntax that is easier to define programmatically.
  xwpp::chart_series_set_categories(series2, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
  xwpp::chart_series_set_values(series2, "Sheet1", 1, 2, 6, 2);     // "=Sheet1!$C$2:$C$7"
  xwpp::chart_series_set_name_range(series2, "Sheet1", 0, 2);       // "=Sheet1!$C$1"

  // Add a chart title and some axis labels.
  chart1.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart1.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart1.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart1.set_style(10);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E2"), &chart1);

  // Chart 2. Create a stacked bar chart.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::LINE_STACKED);

  // Add the first series to the chart.
  xwpp::chart_series_t& series3 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart2.series_set_name(series3, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series4 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart2.series_set_name(series4, "=Sheet1!$C$1");

  // Add a chart title and some axis labels.
  chart2.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart2.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart2.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart2.set_style(12);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E18"), &chart2);

  // Chart 3. Create a percent stacked bar chart.
  xwpp::chart_t& chart3 = workbook.add_chart(xwpp::chart_type_t::LINE_STACKED_PERCENT);

  // Add the first series to the chart.
  xwpp::chart_series_t& series5 = chart3.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart3.series_set_name(series5, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series6 = chart3.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart3.series_set_name(series6, "=Sheet1!$C$1");

  // Add a chart title and some axis labels.
  chart3.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart3.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart3.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart3.set_style(13);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E34"), &chart3);

  workbook.save("chart_line.xlsx");
}
