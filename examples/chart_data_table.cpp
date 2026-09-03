/*
 * An example of creating Excel column charts with data tables using the
 * Xlsxwriter++ library.
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
  const std::array<std::array<uint8_t, 3>, 6> data{
    {{2, 10, 30}, {3, 40, 60}, {4, 50, 70}, {5, 20, 50}, {6, 10, 40}, {7, 50, 30}}
  };

  worksheet.write_string(CELL("A1"), "Number", bold);
  worksheet.write_string(CELL("B1"), "Batch 1", bold);
  worksheet.write_string(CELL("C1"), "Batch 2", bold);

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
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
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a data series to the chart.
    xwpp::chart_series_t& series1 = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1".
    chart.series_set_name(series1, "=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined. They
    // can be defined later using the alternative syntax shown below.
    xwpp::chart_series_t& series2 = chart.add_series("", "");

    // Configure the series using a syntax that is easier to define programmatically.
    xwpp::chart_series_set_categories(series2, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
    xwpp::chart_series_set_values(series2, "Sheet1", 1, 2, 6, 2);     // "=Sheet1!$C$2:$C$7"
    xwpp::chart_series_set_name_range(series2, "Sheet1", 0, 2);       // "=Sheet1!$C$1"

    // Add a chart title and some axis labels.
    chart.title_set_name("Chart with Data Table");
    xwpp::chart_axis_set_name(chart.x_axis_, "Test number");
    xwpp::chart_axis_set_name(chart.y_axis_, "Sample length (mm)");

    // Set a default data table on the X-axis.
    chart.set_table();

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E2"), &chart);
  }

  // Chart 2. Example with value and category data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a data series to the chart.
    xwpp::chart_series_t& series1 = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1".
    chart.series_set_name(series1, "=Sheet1!$B$1");

    // Add the second series to the chart.
    xwpp::chart_series_t& series2 = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2".
    chart.series_set_name(series2, "=Sheet1!$C$1");

    // Add a chart title and some axis labels.
    chart.title_set_name("Data Table with legend keys");
    chart_axis_set_name(chart.x_axis_, "Test number");
    chart_axis_set_name(chart.y_axis_, "Sample length (mm)");

    // Set a data table on the X-axis with the legend keys shown.
    chart.set_table();
    chart.set_table_grid(true, true, true, true);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E18"), &chart);
  }

  workbook.save("chart_data_table.xlsx");
}
