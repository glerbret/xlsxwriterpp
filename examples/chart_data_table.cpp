/*
 * An example of creating Excel column charts with data tables using the
 * Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  int row, col;
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

  for(row = 0; row < 6; row++)
  {
    for(col = 0; col < 3; col++)
    {
      worksheet.write_number(row + 1, col, data[row][col]);
    }
  }
}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  write_worksheet_data(worksheet, bold);

  // Chart 1. Example with standard data labels.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Add a data series to the chart.
  xwpp::chart_series_t& series1 = chart1.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart1.series_set_name(series1, "=Sheet1!$B$1");

  // Add a second series but leave the categories and values undefined. They
  // can be defined later using the alternative syntax shown below.
  xwpp::chart_series_t& series2 = chart1.add_series("", "");

  // Configure the series using a syntax that is easier to define programmatically.
  xwpp::chart_series_set_categories(series2, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
  xwpp::chart_series_set_values(series2, "Sheet1", 1, 2, 6, 2);     // "=Sheet1!$C$2:$C$7"
  xwpp::chart_series_set_name_range(series2, "Sheet1", 0, 2);       // "=Sheet1!$C$1"

  // Add a chart title and some axis labels.
  chart1.title_set_name("Chart with Data Table");
  xwpp::chart_axis_set_name(chart1.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart1.y_axis_, "Sample length (mm)");

  // Set a default data table on the X-axis.
  chart1.set_table();

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E2"), &chart1);

  // Chart 2. Example with value and category data labels.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Add a data series to the chart.
  xwpp::chart_series_t& series3 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart2.series_set_name(series3, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series4 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart2.series_set_name(series4, "=Sheet1!$C$1");

  // Add a chart title and some axis labels.
  chart2.title_set_name("Data Table with legend keys");
  chart_axis_set_name(chart2.x_axis_, "Test number");
  chart_axis_set_name(chart2.y_axis_, "Sample length (mm)");

  // Set a data table on the X-axis with the legend keys shown.
  chart2.set_table();
  chart2.set_table_grid(true, true, true, true);

  // Turn off the legend.
  chart2.legend_set_position(xwpp::chart_legend_position_t::NONE);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E18"), &chart2);

  workbook.save("chart_data_table.xlsx");
}
