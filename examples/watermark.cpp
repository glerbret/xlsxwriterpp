/*
 * An example of adding a worksheet watermark image using Xlsxwriter++. This
 * is based on the method of putting an image in the worksheet header as
 * suggested in the Microsoft documentation:
 * https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Set a worksheet header with the watermark image.
  const xwpp::header_footer_options_t header_options{.image_center_ = "watermark.png"};
  worksheet.set_header("&C&G", header_options);

  workbook.save("watermark.xlsx");
}
