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

  worksheet.write_number(0, 0, 2);
  worksheet.write_number(1, 0, 5);
  worksheet.write_number(2, 0, 4);

  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$A$1:$A$3");

  const xwpp::chart_fill_t fill1{.color_ = xwpp::color_t(0xFF0000)};
  const xwpp::chart_fill_t fill2{.color_ = xwpp::color_t(0xCC0000)};
  const xwpp::chart_fill_t fill3{.color_ = xwpp::color_t(0x990000)};
  const std::vector<xwpp::chart_point_t> points{{.fill_ = fill1}, {.fill_ = fill2}, {.fill_ = fill3}};

  series_set_points(series, points);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_points03.xlsx");
}
