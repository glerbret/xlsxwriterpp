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
  const xwpp::format_t* bold = workbook.format_builder().bold().build();

  // Some chart positioning options.
  xwpp::chart_options_t options{.x_offset = 25, .y_offset = 10};

  // Write some data for the chart.
  worksheet.write(0, 0, "Number", bold);
  worksheet.write(1, 0, 2);
  worksheet.write(2, 0, 3);
  worksheet.write(3, 0, 4);
  worksheet.write(4, 0, 5);
  worksheet.write(5, 0, 6);
  worksheet.write(6, 0, 7);

  worksheet.write(0, 1, "Data", bold);
  worksheet.write(1, 1, 20);
  worksheet.write(2, 1, 10);
  worksheet.write(3, 1, 20);
  worksheet.write(4, 1, 30);
  worksheet.write(5, 1, 40);
  worksheet.write(6, 1, 30);

  worksheet.write(0, 2, "Text", bold);
  worksheet.write(1, 2, "Jan");
  worksheet.write(2, 2, "Feb");
  worksheet.write(3, 2, "Mar");
  worksheet.write(4, 2, "Apr");
  worksheet.write(5, 2, "May");
  worksheet.write(6, 2, "Jun");

  // Chart 1. Example with standard data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Chart with standard data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D2", &chart, options);
  }

  // Chart 2. Example with value and category data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Category and Value data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Turn on Value and Category labels.
    series.set_labels_options(false, true, true);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D18", &chart, options);
  }

  // Chart 3. Example with standard data labels with different font.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Data labels with user defined font");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    const xwpp::chart_font_t font{.bold = true, .rotation = -30, .color = xwpp::color_t::red()};
    xwpp::chart_series_set_labels_font(series, font);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D34", &chart, options);
  }

  // Chart 4. Example with standard data labels and formatting.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Data labels with formatting");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Set the border/line and fill for the data labels.
    const xwpp::chart_line_t line{.color = xwpp::color_t::red()};
    const xwpp::chart_fill_t fill{.color = xwpp::color_t::yellow()};

    xwpp::chart_series_set_labels_line(series, line);
    xwpp::chart_series_set_labels_fill(series, fill);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D50", &chart, options);
  }

  // Chart 5.Example with custom string data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Chart with custom string data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels{
      {.value = "Amy"}, {.value = "Bea"}, {.value = "Eva"}, {.value = "Fay"}, {.value = "Liv"}, {.value = "Una"},
    };

    // Set the custom labels.
    series.set_labels_custom(data_labels);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D66", &chart, options);
  }

  // Chart 6. Example with custom data labels from cells.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Chart with custom data labels from cells");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels{
      {.value = "=Sheet1!$C$2"}, {.value = "=Sheet1!$C$3"}, {.value = "=Sheet1!$C$4"},
      {.value = "=Sheet1!$C$5"}, {.value = "=Sheet1!$C$6"}, {.value = "=Sheet1!$C$7"},
    };

    // Set the custom labels.
    series.set_labels_custom(data_labels);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D82", &chart, options);
  }

  // Chart 7. Example with custom and default data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Mixed custom and default data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    const xwpp::chart_font_t font{.color = xwpp::color_t::red()};

    // Add the series data labels.
    series.set_labels();

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels{
      {.value = "=Sheet1!$C$2", .font = font},
      {},
      {.value = "=Sheet1!$C$4", .font = font},
      {.value = "=Sheet1!$C$5", .font = font},
    };

    // Set the custom labels.
    series.set_labels_custom(data_labels);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D98", &chart, options);
  }

  // Chart 8. Example with deleted/hidden custom data labels.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Chart with deleted data labels");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Create some custom labels.
    const xwpp::chart_data_label_t hide{.hide = true};
    const xwpp::chart_data_label_t keep{.hide = false};
    const std::vector<xwpp::chart_data_label_t> data_labels{
      hide, keep, hide, hide, keep, hide,
    };

    // Set the custom labels.
    series.set_labels_custom(data_labels);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D114", &chart, options);
  }

  // Chart 9.Example with custom string data labels and formatting.
  {
    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);

    // Add a chart title.
    chart.set_title("Chart with custom labels and formatting");

    // Add a data series to the chart.
    xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.set_labels();

    // Set the border/line and fill for the data labels.
    xwpp::chart_line_t line2{.color = xwpp::color_t::red()};
    xwpp::chart_fill_t fill2{.color = xwpp::color_t::yellow()};
    xwpp::chart_line_t line3{.color = xwpp::color_t::blue()};
    xwpp::chart_fill_t fill3{.color = xwpp::color_t::green()};

    // Set the default formatting for the data labels in the series.
    xwpp::chart_series_set_labels_line(series, line2);
    xwpp::chart_series_set_labels_fill(series, fill2);

    // Create some custom labels.
    const std::vector<xwpp::chart_data_label_t> data_labels{
      {.value = "Amy", .line = line3},
      {.value = "Bea"},
      {.value = "Eva"},
      {.value = "Fay"},
      {.value = "Liv"},
      {.value = "Una", .fill = fill3},
    };

    // Set the custom labels.
    series.set_labels_custom(data_labels);

    // Turn off the legend.
    chart.set_legend_position(xwpp::chart_legend_position_t::NONE);

    // Insert the chart into the worksheet.
    worksheet.insert_chart("D130", &chart, options);
  }

  workbook.save("chart_data_labels.xlsx");
}
