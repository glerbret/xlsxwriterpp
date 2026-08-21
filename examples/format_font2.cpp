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

  // Write some formatted strings.
  {
    xwpp::format_t* format = workbook.add_format();
    format->set_bold();

    worksheet.write_string(0, 0, "This is bold", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_italic();

    worksheet.write_string(1, 0, "This is italic", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_bold();
    format->set_italic();

    worksheet.write_string(2, 0, "Bold and italic", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_color(xwpp::color_t::RED);

    worksheet.write_string(3, 0, "Red", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_underline(xwpp::format_underlines_t::SINGLE);

    worksheet.write_string(4, 0, "Underline", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_name("Times New Roman");

    worksheet.write_string(5, 0, "Times New Roman", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_size(24.);

    worksheet.write_string(6, 0, "Font size 24", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_strikeout();

    worksheet.write_string(7, 0, "Strikeout", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_script(xwpp::format_scripts_t::SUPERSCRIPT);

    worksheet.write_string(8, 0, "Superscript", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_script(xwpp::format_scripts_t::SUBSCRIPT);

    worksheet.write_string(9, 0, "Subscript", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_outline();

    worksheet.write_string(10, 0, "Outline", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_shadow();

    worksheet.write_string(11, 0, "Shadow", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_condense();

    worksheet.write_string(12, 0, "Condensed", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_extend();

    worksheet.write_string(13, 0, "Extended", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_rotation(270);

    worksheet.write_string(14, 0, "Vertical text", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_rotation(45);

    worksheet.write_string(15, 0, "With 45°", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_indent(1);

    worksheet.write_string(16, 0, "First level indent", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_indent(2);

    worksheet.write_string(17, 0, "Second level indent", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_shrink();

    worksheet.write_string(18, 0, "Shrink long long long long long long long long long text", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_reading_order(1);

    worksheet.write_string(19, 0, "Reading order 1", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_reading_order(2);

    worksheet.write_string(20, 0, "Reading order 2", format);
  }

  workbook.save("format_font2.xlsx");
}
