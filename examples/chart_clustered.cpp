/*
 * An example of a clustered category chart using the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* bold)
{
  worksheet.write_string(0, 0, "Types", bold);
  worksheet.write_string(1, 0, "Type 1");
  worksheet.write_string(4, 0, "Type 2");

  worksheet.write_string(0, 1, "Sub Type", bold);
  worksheet.write_string(1, 1, "Sub Type A");
  worksheet.write_string(2, 1, "Sub Type B");
  worksheet.write_string(3, 1, "Sub Type C");
  worksheet.write_string(4, 1, "Sub Type D");
  worksheet.write_string(5, 1, "Sub Type E");

  worksheet.write_string(0, 2, "Value 1", bold);
  worksheet.write_number(1, 2, 5000);
  worksheet.write_number(2, 2, 2000);
  worksheet.write_number(3, 2, 250);
  worksheet.write_number(4, 2, 6000);
  worksheet.write_number(5, 2, 500);

  worksheet.write_string(0, 3, "Value 2", bold);
  worksheet.write_number(1, 3, 8000);
  worksheet.write_number(2, 3, 3000);
  worksheet.write_number(3, 3, 1000);
  worksheet.write_number(4, 3, 6000);
  worksheet.write_number(5, 3, 300);

  worksheet.write_string(0, 4, "Value 3", bold);
  worksheet.write_number(1, 4, 6000);
  worksheet.write_number(2, 4, 4000);
  worksheet.write_number(3, 4, 2000);
  worksheet.write_number(4, 4, 6500);
  worksheet.write_number(5, 4, 200);
}

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::COLUMN);

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  write_worksheet_data(worksheet, bold);

  /*
   * Configure the series. Note, that the categories are 2D ranges (from
   * column A to column B). This creates the clusters. The series are shown
   * as formula strings for clarity but you can also use variables with the
   * chart_series_set_categories() and chart_series_set_values()
   * functions. See the docs.
   */
  chart.add_series("=Sheet1!$A$2:$B$6", "=Sheet1!$C$2:$C$6");
  chart.add_series("=Sheet1!$A$2:$B$6", "=Sheet1!$D$2:$D$6");
  chart.add_series("=Sheet1!$A$2:$B$6", "=Sheet1!$E$2:$E$6");

  // Set an Excel chart style.
  chart.set_style(37);

  // Turn off the legend.
  chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("G3"), &chart);

  workbook.save("chart_clustered2.xlsx");
}
