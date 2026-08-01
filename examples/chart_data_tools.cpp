/*
 * A demo of an various Excel chart data tools that are available via a
 * Xlsxwriter++ chart.
 *
 * These include Drop Lines and High-Low Lines.
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
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Write some data for the chart.
  write_worksheet_data(worksheet, bold);

  // Chart 1. Example with High Low Lines.
  xwpp::chart_t& chart1 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart1.title_set_name("Chart with High-Low Lines");

  // Add the first series to the chart.
  chart1.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart1.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add high-low lines to the chart.
  chart1.set_high_low_lines(std::nullopt);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E2"), &chart1);

  // Chart 2. Example with Drop Lines.
  xwpp::chart_t& chart2 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart2.title_set_name("Chart with Drop Lines");

  // Add the first series to the chart.
  chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart2.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add drop lines to the chart.
  chart2.set_drop_lines(std::nullopt);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E18"), &chart2);

  // Chart 3. Example with Up-Down bars.
  xwpp::chart_t& chart3 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart3.title_set_name("Chart with Up-Down bars");

  // Add the first series to the chart.
  chart3.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart3.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add Up-Down bars to the chart.
  chart3.set_up_down_bars();

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E34"), &chart3);

  // Chart 4. Example with Up-Down bars with formatting.
  xwpp::chart_t& chart4 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart4.title_set_name("Chart with Up-Down bars");

  // Add the first series to the chart.
  chart4.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart4.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add Up-Down bars to the chart, with formatting.
  xwpp::chart_line_t line      = {.color_ = xwpp::color_t::BLACK};
  xwpp::chart_fill_t up_fill   = {.color_ = static_cast<xwpp::color_t>(0x00B050)};
  xwpp::chart_fill_t down_fill = {.color_ = xwpp::color_t::RED};

  chart4.set_up_down_bars_format(line, up_fill, line, down_fill);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E50"), &chart4);

  // Chart 5. Example with Markers and data labels.
  xwpp::chart_t& chart5 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart5.title_set_name("Chart with Data Labels and Markers");

  // Add the first series to the chart.
  xwpp::chart_series_t& series1 = chart5.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart5.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add series markers.
  chart5.series_set_marker_type(series1, xwpp::chart_marker_type_t::CIRCLE);

  // Add series data labels.
  xwpp::chart_series_set_labels(series1);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E66"), &chart5);

  // Chart 6. Example with Error Bars.
  xwpp::chart_t& chart6 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart6.title_set_name("Chart with Error Bars");

  // Add the first series to the chart.
  xwpp::chart_series_t& series2 = chart6.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart6.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add error bars to show Standard Error.
  chart6.series_set_error_bars(series2.y_error_bars_, xwpp::chart_error_bar_type_t::STD_ERROR, 0);

  // Add series data labels.
  xwpp::chart_series_set_labels(series2);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E82"), &chart6);

  // Chart 7. Example with a trendline
  xwpp::chart_t& chart7 = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Add a chart title.
  chart7.title_set_name("Chart with a Trendline");

  // Add the first series to the chart.
  xwpp::chart_series_t& series3 = chart7.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
  chart7.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

  // Add a polynomial trendline.
  xwpp::chart_line_t poly_line = {.color_     = xwpp::color_t::GRAY,
                                  .dash_type_ = xwpp::chart_line_dash_type_t::DASH_LONG_DASH};

  xwpp::series_set_trendline(series3, xwpp::chart_trendline_type_t::POLY, 3);
  xwpp::series_set_trendline_line(series3, poly_line);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("E98"), &chart7);

  workbook.save("chart_data_tools.xlsx");
}
