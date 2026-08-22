/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet   = workbook.add_worksheet();
  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
  xwpp::chart_t& chart           = workbook.add_chart(xwpp::chart_type_t::BAR);

  // For testing, copy the randomly generated axis ids in the target file.
  chart.set_axis_ids(61297792, 61299328);

  const std::array<std::array<uint8_t, 3>, 5> data{
    {{1, 2, 3}, {2, 4, 6}, {3, 6, 9}, {4, 8, 12}, {5, 10, 15}}
  };

  for(xwpp::row_num_t row_num = 0; const auto& row: data)
  {
    for(xwpp::col_num_t col_num = 0; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }

  chart.add_series("", "=Sheet1!$A$1:$A$5");
  chart.add_series("", "=Sheet1!$B$1:$B$5");
  chart.add_series("", "=Sheet1!$C$1:$C$5");

  chartsheet.set_margins(0.511811023622047, 0.511811023622047, 0.551181102362204, 0.944881889763779);

  chartsheet.set_paper(9);
  chartsheet.set_portrait();

  const xwpp::header_footer_options_t header_options{.margin_ = 0.118110236220472};
  chartsheet.set_header("&CPage &P", header_options);
  chartsheet.set_footer("&C&A", header_options);

  // For testing.
  chartsheet.set_dpi(200, 200);

  chartsheet.set_chart(&chart);

  workbook.save("test_chartsheet08.xlsx");
}
