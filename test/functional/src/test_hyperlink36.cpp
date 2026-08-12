/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  xwpp::chart_t& chart         = workbook.add_chart(xwpp::chart_type_t::PIE);

  const xwpp::image_options_t options{.url_ = "https://github.com/jmcnamara"};

  worksheet.write_number(CELL("A1"), 1);
  worksheet.write_number(CELL("A2"), 2);

  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  chart.add_series("", "=Sheet1!$A$1:$A$2");

  worksheet.insert_chart(CELL("E12"), &chart);

  workbook.save("test_hyperlink36.xlsx");
}
