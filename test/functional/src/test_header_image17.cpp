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

  const xwpp::header_footer_options_t header_options1{
    .image_left_ = "images/red.jpg", .image_center_ = "images/blue.jpg", .image_right_ = "images/red.jpg"};

  const xwpp::header_footer_options_t header_options2{
    .image_left_ = "images/blue.jpg", .image_center_ = "images/red.jpg", .image_right_ = "images/blue.jpg"};

  worksheet.set_header("&L&G&C&G&R&G", header_options1);
  worksheet.set_footer("&L&G&C&G&R&G", header_options2);

  workbook.save("test_header_image17.xlsx");
}
