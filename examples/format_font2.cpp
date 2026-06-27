/*
 * Example of writing some data with font formatting to a simple Excel
 * file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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

  xwpp::format_t* format15 = workbook.add_format();
  format15->set_rotation(270);

  xwpp::format_t* format16 = workbook.add_format();
  format16->set_rotation(45);

  xwpp::format_t* format17 = workbook.add_format();
  format17->set_indent(1);

  xwpp::format_t* format18 = workbook.add_format();
  format18->set_indent(2);

  xwpp::format_t* format19 = workbook.add_format();
  format19->set_shrink();

  xwpp::format_t* format20 = workbook.add_format();
  format20->set_reading_order(1);

  xwpp::format_t* format21 = workbook.add_format();
  format21->set_reading_order(2);

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
  worksheet.write_string(14, 0, "Vertical text", format15);
  worksheet.write_string(15, 0, "With 45°", format16);
  worksheet.write_string(16, 0, "First level indent", format17);
  worksheet.write_string(17, 0, "Second level indent", format18);
  worksheet.write_string(18, 0, "Shrink long long long long long long long long long text", format19);
  worksheet.write_string(19, 0, "Reading order 1", format20);
  worksheet.write_string(20, 0, "Reading order 2", format21);

  workbook.save("format_font2.xlsx");
}
