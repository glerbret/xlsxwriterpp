/*
 * A demo of an various Excel chart data label features that are available via
 * a Xlsxwriter++ chart.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a bold format to use to highlight the header cells.
  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  // Some chart positioning options.
  xwpp::chart_options_t options = {.x_offset_ = 25, .y_offset_ = 10};

  // Write some data for the chart.
  worksheet.write_string(0, 0, "Number", bold);
  worksheet.write_number(1, 0, 2);
  worksheet.write_number(2, 0, 3);
  worksheet.write_number(3, 0, 4);
  worksheet.write_number(4, 0, 5);
  worksheet.write_number(5, 0, 6);
  worksheet.write_number(6, 0, 7);

  worksheet.write_string(0, 1, "Data", bold);
  worksheet.write_number(1, 1, 20);
  worksheet.write_number(2, 1, 10);
  worksheet.write_number(3, 1, 20);
  worksheet.write_number(4, 1, 30);
  worksheet.write_number(5, 1, 40);
  worksheet.write_number(6, 1, 30);

  worksheet.write_string(0, 2, "Text", bold);
  worksheet.write_string(1, 2, "Jan");
  worksheet.write_string(2, 2, "Feb");
  worksheet.write_string(3, 2, "Mar");
  worksheet.write_string(4, 2, "Apr");
  worksheet.write_string(5, 2, "May");
  worksheet.write_string(6, 2, "Jun");

  // Chart 1. Example with standard data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Chart with standard data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D2"), &chart, options);
  }

  // Chart 2. Example with value and category data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Category and Value data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Turn on Value and Category labels.
    xwpp::chart_series_set_labels_options(series, false, true, true);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D18"), &chart, options);
  }

  // Chart 3. Example with standard data labels with different font.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Data labels with user defined font");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    const xwpp::chart_font_t font{.bold_ = true, .rotation_ = -30, .color_ = xwpp::color_t::RED};
    xwpp::chart_series_set_labels_font(series, font);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D34"), &chart, options);
  }

  // Chart 4. Example with standard data labels and formatting.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Data labels with formatting");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Set the border/line and fill for the data labels.
    const xwpp::chart_line_t line{.color_ = xwpp::color_t::RED};
    const xwpp::chart_fill_t fill{.color_ = xwpp::color_t::YELLOW};

    xwpp::chart_series_set_labels_line(series, line);
    xwpp::chart_series_set_labels_fill(series, fill);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D50"), &chart, options);
  }

  // Chart 5.Example with custom string data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Chart with custom string data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels = {
      {.value_ = "Amy"}, {.value_ = "Bea"}, {.value_ = "Eva"}, {.value_ = "Fay"}, {.value_ = "Liv"}, {.value_ = "Una"},
    };

    // Set the custom labels.
    chart_series_set_labels_custom(series, data_labels);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D66"), &chart, options);
  }

  // Chart 6. Example with custom data labels from cells.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Chart with custom data labels from cells");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels = {
      {.value_ = "=Sheet1!$C$2"}, {.value_ = "=Sheet1!$C$3"}, {.value_ = "=Sheet1!$C$4"},
      {.value_ = "=Sheet1!$C$5"}, {.value_ = "=Sheet1!$C$6"}, {.value_ = "=Sheet1!$C$7"},
    };

    // Set the custom labels.
    chart_series_set_labels_custom(series, data_labels);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D82"), &chart, options);
  }

  // Chart 7. Example with custom and default data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Mixed custom and default data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    const xwpp::chart_font_t font{.color_ = xwpp::color_t::RED};

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels = {
      {.value_ = "=Sheet1!$C$2", .font_ = font},
      {},
      {.value_ = "=Sheet1!$C$4", .font_ = font},
      {.value_ = "=Sheet1!$C$5", .font_ = font},
    };

    // Set the custom labels.
    chart_series_set_labels_custom(series, data_labels);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D98"), &chart, options);
  }

  // Chart 8. Example with deleted/hidden custom data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Chart with deleted data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Create some custom labels.
    const xwpp::chart_data_label_t hide{.hide_ = true};
    const xwpp::chart_data_label_t keep{.hide_ = false};
    const std::vector<xwpp::chart_data_label_t> data_labels{
      hide, keep, hide, hide, keep, hide,
    };

    // Set the custom labels.
    chart_series_set_labels_custom(series, data_labels);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D114"), &chart, options);
  }

  // Chart 9.Example with custom string data labels and formatting.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.title_set_name("Chart with custom labels and formatting");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    xwpp::chart_series_set_labels(series);

    // Set the border/line and fill for the data labels.
    xwpp::chart_line_t line2{.color_ = xwpp::color_t::RED};
    xwpp::chart_fill_t fill2{.color_ = xwpp::color_t::YELLOW};
    xwpp::chart_line_t line3{.color_ = xwpp::color_t::BLUE};
    xwpp::chart_fill_t fill3{.color_ = xwpp::color_t::GREEN};

    // Set the default formatting for the data labels in the series.
    xwpp::chart_series_set_labels_line(series, line2);
    xwpp::chart_series_set_labels_fill(series, fill2);

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels{
      {.value_ = "Amy", .line_ = line3},
      {.value_ = "Bea"},
      {.value_ = "Eva"},
      {.value_ = "Fay"},
      {.value_ = "Liv"},
      {.value_ = "Una", .fill_ = fill3},
    };

    // Set the custom labels.
    chart_series_set_labels_custom(series, data_labels);

    // Turn off the legend.
    chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart(CELL("D130"), &chart, options);
  }

  workbook.save("chart_data_labels.xlsx");
}
