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

  // Chart 1. Example with High Low Lines.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with High-Low Lines");

    // Add the first series to the chart.
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add high-low lines to the chart.
    chart.set_high_low_lines(std::nullopt);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E2"), &chart);
  }

  // Chart 2. Example with Drop Lines.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with Drop Lines");

    // Add the first series to the chart.
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add drop lines to the chart.
    chart.set_drop_lines(std::nullopt);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E18"), &chart);
  }

  // Chart 3. Example with Up-Down bars.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with Up-Down bars");

    // Add the first series to the chart.
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart.
    chart.set_up_down_bars();

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E34"), &chart);
  }

  // Chart 4. Example with Up-Down bars with formatting.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with Up-Down bars");

    // Add the first series to the chart.
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart, with formatting.
    xwpp::chart_line_t line{.color_ = xwpp::color_t::BLACK};
    xwpp::chart_fill_t up_fill{.color_ = static_cast<xwpp::color_t>(0x00B050)};
    xwpp::chart_fill_t down_fill{.color_ = xwpp::color_t::RED};

    chart.set_up_down_bars_format(line, up_fill, line, down_fill);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E50"), &chart);
  }

  // Chart 5. Example with Markers and data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with Data Labels and Markers");

    // Add the first series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add series markers.
    chart.series_set_marker_type(series, xwpp::chart_marker_type_t::CIRCLE);

    // Add series data labels.
    xwpp::chart_series_set_labels(series);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E66"), &chart);
  }

  // Chart 6. Example with Error Bars.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with Error Bars");

    // Add the first series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add error bars to show Standard Error.
    chart.series_set_error_bars(series.y_error_bars_, xwpp::chart_error_bar_type_t::STD_ERROR, 0);

    // Add series data labels.
    xwpp::chart_series_set_labels(series);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E82"), &chart);
  }

  // Chart 7. Example with a trendline
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

    // Add a chart title.
    chart.title_set_name("Chart with a Trendline");

    // Add the first series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add a polynomial trendline.
    xwpp::chart_line_t poly_line{.color_ = xwpp::color_t::GRAY, .dash_type_ = xwpp::chart_line_dash_type_t::LONG_DASH};

    xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::POLY, 3);
    xwpp::series_set_trendline_line(series, poly_line);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("E98"), &chart);
  }

  workbook.save("chart_data_tools.xlsx");
}
