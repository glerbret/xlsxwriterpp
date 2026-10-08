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
    const xwpp::format_t* format = workbook.format_builder().bold().build();
    worksheet.write(0, 0, "This is bold", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().italic().build();
    worksheet.write(1, 0, "This is italic", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().bold().italic().build();
    worksheet.write(2, 0, "Bold and italic", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_color(xwpp::color_t::red()).build();
    worksheet.write(3, 0, "Red", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().underline(xwpp::format_underlines_t::SINGLE).build();
    worksheet.write(4, 0, "Underline", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_name("Times New Roman").build();
    worksheet.write(5, 0, "Times New Roman", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_size(24.).build();
    worksheet.write(6, 0, "Font size 24", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_strikeout().build();
    worksheet.write(7, 0, "Strikeout", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_script(xwpp::format_scripts_t::SUPERSCRIPT).build();
    worksheet.write(8, 0, "Superscript", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_script(xwpp::format_scripts_t::SUBSCRIPT).build();
    worksheet.write(9, 0, "Subscript", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_outline().build();
    worksheet.write(10, 0, "Outline", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_shadow().build();
    worksheet.write(11, 0, "Shadow", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_condense().build();
    worksheet.write(12, 0, "Condensed", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().font_extend().build();
    worksheet.write(13, 0, "Extended", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().rotation(270).build();
    worksheet.write(14, 0, "Vertical text", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().rotation(45).build();
    worksheet.write(15, 0, "With 45°", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().indent(1).build();
    worksheet.write(16, 0, "First level indent", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().indent(2).build();
    worksheet.write(17, 0, "Second level indent", format);
  }

  {
    const xwpp::format_t* format = workbook.format_builder().shrink().build();
    worksheet.write(18, 0, "Shrink long long long long long long long long long text", format);
  }

  {
    const xwpp::format_t* format =
      workbook.format_builder().reading_order(xwpp::format_reading_order_t::LEFT_TO_RIGHT).build();
    worksheet.write(19, 0, "Reading order 1", format);
  }

  {
    const xwpp::format_t* format =
      workbook.format_builder().reading_order(xwpp::format_reading_order_t::RIGHT_TO_LEFT).build();
    worksheet.write(20, 0, "Reading order 2", format);
  }

  workbook.save("format_font2.xlsx");
}
