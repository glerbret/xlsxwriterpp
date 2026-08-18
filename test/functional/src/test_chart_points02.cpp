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
  worksheet.write_number(3, 0, 1);
  worksheet.write_number(4, 0, 7);
  worksheet.write_number(5, 0, 4);

  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$A$1:$A$6");

  const xwpp::chart_line_t line{.color_     = xwpp::color_t::RED,
                                .dash_type_ = xwpp::chart_line_dash_type_t::DASH_SQUARE_DOT};
  const xwpp::chart_fill_t fill{.color_ = xwpp::color_t::YELLOW};
  const std::vector<xwpp::chart_point_t> points{{}, {.line_ = line}, {}, {.fill_ = fill}};

  series_set_points(series, points);

  worksheet.insert_chart(CELL("E9"), &chart);

  workbook.save("test_chart_points02.xlsx");
}
