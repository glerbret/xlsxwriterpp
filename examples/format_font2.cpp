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

    worksheet.write(0, 0, "This is bold", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_italic();

    worksheet.write(1, 0, "This is italic", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_bold();
    format->set_italic();

    worksheet.write(2, 0, "Bold and italic", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_color(xwpp::color_t::red());

    worksheet.write(3, 0, "Red", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_underline(xwpp::format_underlines_t::SINGLE);

    worksheet.write(4, 0, "Underline", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_name("Times New Roman");

    worksheet.write(5, 0, "Times New Roman", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_size(24.);

    worksheet.write(6, 0, "Font size 24", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_strikeout();

    worksheet.write(7, 0, "Strikeout", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_script(xwpp::format_scripts_t::SUPERSCRIPT);

    worksheet.write(8, 0, "Superscript", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_script(xwpp::format_scripts_t::SUBSCRIPT);

    worksheet.write(9, 0, "Subscript", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_outline();

    worksheet.write(10, 0, "Outline", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_shadow();

    worksheet.write(11, 0, "Shadow", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_condense();

    worksheet.write(12, 0, "Condensed", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_font_extend();

    worksheet.write(13, 0, "Extended", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_rotation(270);

    worksheet.write(14, 0, "Vertical text", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_rotation(45);

    worksheet.write(15, 0, "With 45°", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_indent(1);

    worksheet.write(16, 0, "First level indent", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_indent(2);

    worksheet.write(17, 0, "Second level indent", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_shrink();

    worksheet.write(18, 0, "Shrink long long long long long long long long long text", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_reading_order(1);

    worksheet.write(19, 0, "Reading order 1", format);
  }

  {
    xwpp::format_t* format = workbook.add_format();
    format->set_reading_order(2);

    worksheet.write(20, 0, "Reading order 2", format);
  }

  workbook.save("format_font2.xlsx");
}
