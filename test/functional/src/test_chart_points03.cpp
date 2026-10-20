/*
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
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

  worksheet.write(0, 0, 2);
  worksheet.write(1, 0, 5);
  worksheet.write(2, 0, 4);

  const xwpp::chart_fill_t fill1{.color = xwpp::color_t{0xFF0000}};
  const xwpp::chart_fill_t fill2{.color = xwpp::color_t{0xCC0000}};
  const xwpp::chart_fill_t fill3{.color = xwpp::color_t{0x990000}};
  const std::vector<xwpp::chart_point_t> points{{.fill = fill1}, {.fill = fill2}, {.fill = fill3}};
  chart.add_series("=Sheet1!$A$1:$A$3").set_points(points);

  worksheet.insert_chart("E9", &chart);

  workbook.save("test_chart_points03.xlsx");
}
