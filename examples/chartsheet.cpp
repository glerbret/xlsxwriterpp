/*
 * An example of creating an Excel chartsheet using the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
  xwpp::worksheet_t& worksheet   = workbook.add_worksheet();
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  write_worksheet_data(worksheet, bold);

  // Create a bar chart.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::BAR);

  // Add the first series to the chart.
  xwpp::chart_series_t& series1 = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart.series_set_name(series1, "=Sheet1!$B$1");

  // Add a second series but leave the categories and values undefined. They
  // can be defined later using the alternative syntax shown below.  */
  xwpp::chart_series_t& series2 = chart.add_series("", "");

  // Configure the series using a syntax that is easier to define programmatically.
  xwpp::chart_series_set_categories(series2, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
  xwpp::chart_series_set_values(series2, "Sheet1", 1, 2, 6, 2);     // "=Sheet1!$C$2:$C$7"
  xwpp::chart_series_set_name_range(series2, "Sheet1", 0, 2);       // "=Sheet1!$C$1"

  // Add a chart title and some axis labels.
  chart.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart.set_style(11);

  // Add the chart to the chartsheet.
  chartsheet.set_chart(&chart);

  // Display the chartsheet as the active sheet when the workbook is opened.
  chartsheet.activate();

  workbook.save("chartsheet.xlsx");
}
