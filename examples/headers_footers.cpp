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
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet("Simple");

  worksheet1.set_header("&CHere is some centered text.");
  worksheet1.set_footer("&LHere is some left aligned text.");

  worksheet1.set_column(0, 0, 50);
  worksheet1.write_string(0, 0, preview);

  // A simple example to start
  xwpp::worksheet_t& worksheet2                      = workbook.add_worksheet("Image");
  const xwpp::header_footer_options_t header_options = {.image_left_ = "logo_small.png"};

  worksheet2.set_header("&L&G", header_options);

  worksheet2.set_margins(-1, -1, 1.3, -1);
  worksheet2.set_column(0, 0, 50);
  worksheet2.write_string(0, 0, preview);

  // This is an example of some of the header/footer variables.
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet("Variables");

  worksheet3.set_header("&LPage &P of &N"
                        "&CFilename: &F"
                        "&RSheetname: &A");
  worksheet3.set_footer("&LCurrent date: &D"
                        "&RCurrent time: &T");

  worksheet3.set_column(0, 0, 50);
  worksheet3.write_string(0, 0, preview);

  worksheet3.set_h_pagebreaks({20});
  worksheet3.write_string(20, 0, "Next page");

  // This example shows how to use more than one font.
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet("Mixed fonts");

  worksheet4.set_header(R"(&C&"Courier New,Bold"Hello &"Arial,Italic"World)");
  worksheet4.set_footer(R"(&C&"Symbol"e&"Arial" = mc&X2)");

  worksheet4.set_column(0, 0, 50);
  worksheet4.write_string(0, 0, preview);

  // Example of line wrapping.
  xwpp::worksheet_t& worksheet5 = workbook.add_worksheet("Word wrap");

  worksheet5.set_header("&CHeading 1\nHeading 2");

  worksheet5.set_column(0, 0, 50);
  worksheet5.write_string(0, 0, preview);

  // Example of inserting a literal ampersand &
  xwpp::worksheet_t& worksheet6 = workbook.add_worksheet("Ampersand");

  worksheet6.set_header("&CCuriouser && Curiouser - Attorneys at Law");

  worksheet6.set_column(0, 0, 50);
  worksheet6.write_string(0, 0, preview);

  workbook.save("headers_footers.xlsx");
}
