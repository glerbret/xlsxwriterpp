/*
 * This program shows several examples of how to set up headers and
 * footers with Xlsxwriter++.
 *
 * The control characters used in the header/footer strings are:
 *
 *     Control             Category            Description
 *     =======             ========            ===========
 *     &L                  Justification       Left
 *     &C                                      Center
 *     &R                                      Right
 *
 *     &P                  Information         Page number
 *     &N                                      Total number of pages
 *     &D                                      Date
 *     &T                                      Time
 *     &F                                      File name
 *     &A                                      Worksheet name
 *
 *     &fontsize           Font                Font size
 *     &"font,style"                           Font name and style
 *     &U                                      Single underline
 *     &E                                      Double underline
 *     &S                                      Strikethrough
 *     &X                                      Superscript
 *     &Y                                      Subscript
 *
 *     &G                  Images              Image placeholder
 *
 *     &&                  Miscellaneous       Literal ampersand &
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>

int main()
{
  xwpp::workbook_t workbook;

  const std::string preview = "Select Print Preview to see the header and footer";

  // A simple example to start
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Simple");

    worksheet.set_header("&CHere is some centered text.");
    worksheet.set_footer("&LHere is some left aligned text.");

    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);
  }

  // A simple example to start
  {
    xwpp::worksheet_t& worksheet                       = workbook.add_worksheet("Image");
    const xwpp::header_footer_options_t header_options = {.image_left_ = "logo_small.png"};

    worksheet.set_header("&L&G", header_options);

    worksheet.set_margins(-1, -1, 1.3, -1);
    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);
  }

  // This is an example of some of the header/footer variables.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Variables");

    worksheet.set_header("&LPage &P of &N"
                         "&CFilename: &F"
                         "&RSheetname: &A");
    worksheet.set_footer("&LCurrent date: &D"
                         "&RCurrent time: &T");

    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);

    worksheet.set_h_pagebreaks({20});
    worksheet.write_string(20, 0, "Next page");
  }

  // This example shows how to use more than one font.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Mixed fonts");

    worksheet.set_header(R"(&C&"Courier New,Bold"Hello &"Arial,Italic"World)");
    worksheet.set_footer(R"(&C&"Symbol"e&"Arial" = mc&X2)");

    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);
  }

  // Example of line wrapping.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Word wrap");

    worksheet.set_header("&CHeading 1\nHeading 2");

    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);
  }

  // Example of inserting a literal ampersand &
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Ampersand");

    worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");

    worksheet.set_column(0, 0, 50);
    worksheet.write_string(0, 0, preview);
  }

  workbook.save("headers_footers.xlsx");
}
