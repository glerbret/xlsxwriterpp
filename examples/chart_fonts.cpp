/*
 * An example of a simple Excel chart with user defined fonts using the
 * Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some data for the chart.
  worksheet.write_number(0, 0, 10);
  worksheet.write_number(1, 0, 40);
  worksheet.write_number(2, 0, 50);
  worksheet.write_number(3, 0, 20);
  worksheet.write_number(4, 0, 10);
  worksheet.write_number(5, 0, 50);

  // Create a chart object.
  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);

  // Configure the chart.
  chart.add_series("", "Sheet1!$A$1:$A$6");

  // Create some fonts to use in the chart.
  xwpp::chart_font_t font1{.name_ = "Calibri", .color_ = xwpp::color_t::BLUE};
  xwpp::chart_font_t font2{.name_ = "Courier", .color_ = static_cast<xwpp::color_t>(0x92D050)};
  xwpp::chart_font_t font3{.name_ = "Arial", .color_ = static_cast<xwpp::color_t>(0x00B0F0)};
  xwpp::chart_font_t font4{.name_ = "Century", .color_ = xwpp::color_t::RED};
  xwpp::chart_font_t font5{.rotation_ = -30};
  xwpp::chart_font_t font6{
    .bold_ = true, .italic_ = true, .underline_ = true, .color_ = static_cast<xwpp::color_t>(0x7030A0)};

  // Write the chart title with a font.
  chart.title_set_name("Test Results");
  chart.title_set_name_font(font1);

  // Write the Y axis with a font.
  xwpp::chart_axis_set_name(chart.y_axis_, "Units");
  xwpp::chart_axis_set_name_font(chart.y_axis_, font2);
  xwpp::chart_axis_set_num_font(chart.y_axis_, font3);

  // Write the X axis with a font.
  xwpp::chart_axis_set_name(chart.x_axis_, "Month");
  xwpp::chart_axis_set_name_font(chart.x_axis_, font4);
  xwpp::chart_axis_set_num_font(chart.x_axis_, font5);

  // Display the chart legend at the bottom of the chart.
  chart.legend_set_position(xwpp::chart_legend_position_t::BOTTOM);
  chart.legend_set_font(font6);

  // Insert the chart into the worksheet.
  worksheet.insert_chart(CELL("C1"), &chart);

  workbook.save("chart_fonts.xlsx");
}
