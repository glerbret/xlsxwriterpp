/*
 * An example showing all 48 default chart styles available in Excel 2007
 * using the Xlsxwriter++ library. Note, these styles are not the same as the
 * styles available in Excel 2013.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  const std::vector<std::tuple<xwpp::chart_type_t, std::string>> chart_descs{
    {xwpp::chart_type_t::COLUMN, "Column"},
    {xwpp::chart_type_t::AREA,   "Area"  },
    {xwpp::chart_type_t::LINE,   "Line"  },
    {xwpp::chart_type_t::PIE,    "Pie"   },
  };

  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;

  for(const auto& [chart_type, chart_name]: chart_descs)
  {
    // Add a worksheet for each chart type.
    xwpp::worksheet_t& worksheet = workbook.add_worksheet(chart_name);
    worksheet.set_zoom(30);

    // Create 48 charts, each with a different style.
    uint8_t style_num = 1;
    for(xwpp::row_num_t row_num = 0; row_num < 90; row_num += 15)
    {
      for(xwpp::col_num_t col_num = 0; col_num < 64; col_num += 8)
      {
        xwpp::chart_t& chart = workbook.add_chart(chart_type);
        chart.add_series("", "=Data!$A$1:$A$6");
        chart.title_set_name("Style " + std::to_string(style_num));
        chart.set_style(style_num);
        worksheet.insert_chart(row_num, col_num, &chart);

        style_num++;
      }
    }
  }

  // Create a worksheet with data for the charts.
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Data");
  worksheet.write_number(0, 0, 10);
  worksheet.write_number(1, 0, 40);
  worksheet.write_number(2, 0, 50);
  worksheet.write_number(3, 0, 20);
  worksheet.write_number(4, 0, 10);
  worksheet.write_number(5, 0, 50);

  workbook.save("chart_styles.xlsx");
}
