/*
 * Example of writing some data with font formatting to a simple Excel
 * file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 20);

  // Add some formats.
  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bold();

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_italic();

  xwpp::format_t* format3 = workbook.add_format();
  format3->set_bold();
  format3->set_italic();

  xwpp::format_t* format4 = workbook.add_format();
  format4->set_font_color(xwpp::color_t::RED);

  xwpp::format_t* format5 = workbook.add_format();
  format5->set_underline(xwpp::format_underlines_t::SINGLE);

  xwpp::format_t* format6 = workbook.add_format();
  format6->set_font_name("Times New Roman");

  xwpp::format_t* format7 = workbook.add_format();
  format7->set_font_size(24.);

  xwpp::format_t* format8 = workbook.add_format();
  format8->set_font_strikeout();

  xwpp::format_t* format9 = workbook.add_format();
  format9->set_font_script(xwpp::format_scripts_t::SUPERSCRIPT);

   xwpp::format_t* format10 = workbook.add_format();
  format10->set_font_script(xwpp::format_scripts_t::SUBSCRIPT);

  xwpp::format_t* format11 = workbook.add_format();
  format11->set_font_outline();

  xwpp::format_t* format12 = workbook.add_format();
  format12->set_font_shadow();

  xwpp::format_t* format13 = workbook.add_format();
  format13->set_font_condense();

  xwpp::format_t* format14 = workbook.add_format();
  format14->set_font_extend();

  // Write some formatted strings.
  worksheet.write_string(0, 0, "This is bold", format1);
  worksheet.write_string(1, 0, "This is italic", format2);
  worksheet.write_string(2, 0, "Bold and italic", format3);
  worksheet.write_string(3, 0, "Red", format4);
  worksheet.write_string(4, 0, "Underline", format5);
  worksheet.write_string(5, 0, "Times New Roman", format6);
  worksheet.write_string(6, 0, "Font size 24", format7);
  worksheet.write_string(7, 0, "Strikeout", format8);
  worksheet.write_string(8, 0, "Superscript", format9);
  worksheet.write_string(9, 0, "Subscript", format10);
  worksheet.write_string(10, 0, "Outline", format11);
  worksheet.write_string(11, 0, "Shadow", format12);
  worksheet.write_string(12, 0, "Condensed", format13);
  worksheet.write_string(13, 0, "Extended", format14);

  workbook.save("format_font2.xlsx");
}
