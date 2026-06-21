/*
 * An example of creating Excel scatter charts using the Xlsxwriter++ library.
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

    worksheet.write_string(CELL("A1"), "Number",  bold);
    worksheet.write_string(CELL("B1"), "Batch 1", bold);
    worksheet.write_string(CELL("C1"), "Batch 2", bold);

    for (row = 0; row < 6; row++)
        for (col = 0; col < 3; col++)
            worksheet.write_number(row + 1, col, data[row][col] );
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

  // Chart 1. Create a scatter chart, the default shows points only.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::SCATTER);

  // Add the first series to the chart.
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

  // Add a chart title.
  chart1.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart1.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart1.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart1.set_style(11);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E2"), &chart1);


  // Chart 2. Create a scatter chart with straight lines and markers connecting the points.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::SCATTER_STRAIGHT_WITH_MARKERS);

  // Add the first series to the chart.
  xwpp::chart_series_t& series3 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart2.series_set_name(series3, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series4 = chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart2.series_set_name(series4, "=Sheet1!$C$1");

    /* Add a chart title and some axis labels. */
  chart2.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart2.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart2.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart2.set_style(12);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E18"), &chart2);

  // Chart 3. Create a scatter chart with straight lines connecting the points.
  xwpp::chart_t& chart3 = workbook.add_chart(xwpp::chart_type_t::SCATTER_STRAIGHT);

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

  // Chart 4. Create a scatter chart with smooth lines and markers connecting the points.
  xwpp::chart_t& chart4 = workbook.add_chart(xwpp::chart_type_t::SCATTER_SMOOTH_WITH_MARKERS);

  // Add the first series to the chart.
  xwpp::chart_series_t& series7 = chart4.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart4.series_set_name(series7, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series8 = chart4.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart4.series_set_name(series8, "=Sheet1!$C$1");

   // Add a chart title and some axis labels.
  chart4.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart4.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart4.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart4.set_style(14);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E50"), &chart4);

  // Chart 5. Create a scatter chart with smooth lines connecting the points.
  xwpp::chart_t& chart5 = workbook.add_chart(xwpp::chart_type_t::SCATTER_SMOOTH);

  // Add the first series to the chart.
  xwpp::chart_series_t& series9 = chart5.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

  // Set the name for the series instead of the default "Series 1".
  chart5.series_set_name(series9, "=Sheet1!$B$1");

  // Add the second series to the chart.
  xwpp::chart_series_t& series10 = chart5.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Set the name for the series instead of the default "Series 2".
  chart5.series_set_name(series10, "=Sheet1!$C$1");

  // Add a chart title and some axis labels.
  chart5.title_set_name("Results of sample analysis");
  xwpp::chart_axis_set_name(chart5.x_axis_, "Test number");
  xwpp::chart_axis_set_name(chart5.y_axis_, "Sample length (mm)");

  // Set an Excel chart style.
  chart5.set_style(15);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E66"), &chart5);

  workbook.save("chart_scatter.xlsx");
}
