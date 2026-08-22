/*
 * An example of writing cell comments to a worksheet using Xlsxwriter++.
 *
 * Each of the worksheets demonstrates different features of cell comments.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;

  xwpp::format_t* text_wrap = workbook.add_format();
  text_wrap->set_text_wrap();
  text_wrap->set_align(xwpp::format_alignments_t::VERTICAL_TOP);

  // Example 1. Demonstrates a simple cell comments without formatting.
  // Set up some worksheet formatting.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);

    worksheet.write_string(CELL("C3"), "Hold the mouse over this cell to see the comment.", text_wrap);
    worksheet.write_comment(CELL("C3"), "This is a comment.");
  }

  // Example 2. Demonstrates visible and hidden comments.
  // Set up some worksheet formatting.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);
    worksheet.set_row(2, 50);

    worksheet.write_string(CELL("C3"), "This cell comment is visible.", text_wrap);

    // Use an option to make the comment visible.
    const xwpp::comment_options_t options{.visible_ = xwpp::comment_display_t::VISIBLE};
    worksheet.write_comment(CELL("C3"), "Hello.", options);

    worksheet.write_string(CELL("C6"),
                           "This cell comment isn't visible until you pass "
                           "the mouse over it (the default).",
                           text_wrap);

    worksheet.write_comment(CELL("C6"), "Hello.");
  }

  // Example 3. Demonstrates visible and hidden comments, set at the worksheet level.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);
    worksheet.set_row(5, 50);
    worksheet.set_row(8, 50);

    // Make all comments on the worksheet visible.
    worksheet.show_comments();

    worksheet.write_string(CELL("C3"), "This cell comment is visible, explicitly.", text_wrap);

    const xwpp::comment_options_t options1{.visible_ = xwpp::comment_display_t::VISIBLE};
    worksheet.write_comment(2, 2, "Hello", options1);

    worksheet.write_string(CELL("C6"),
                           "This cell comment is also visible because "
                           "we used worksheet_show_comments().",
                           text_wrap);

    worksheet.write_comment(CELL("C6"), "Hello");

    worksheet.write_string(CELL("C9"), "However, we can still override it locally.", text_wrap);

    const xwpp::comment_options_t options2{.visible_ = xwpp::comment_display_t::HIDDEN};
    worksheet.write_comment(CELL("C9"), "Hello", options2);
  }

  // Example 4. Demonstrates changes to the comment box dimensions.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);
    worksheet.set_row(5, 50);
    worksheet.set_row(8, 50);
    worksheet.set_row(15, 50);
    worksheet.set_row(18, 50);

    worksheet.show_comments();

    worksheet.write_string(CELL("C3"), "This cell comment is default size.", text_wrap);

    worksheet.write_comment(2, 2, "Hello");

    worksheet.write_string(CELL("C6"), "This cell comment is twice as wide.", text_wrap);

    const xwpp::comment_options_t options1{.x_scale_ = 2.0};
    worksheet.write_comment(CELL("C6"), "Hello", options1);

    worksheet.write_string(CELL("C9"), "This cell comment is twice as high.", text_wrap);

    const xwpp::comment_options_t options2{.y_scale_ = 2.0};
    worksheet.write_comment(CELL("C9"), "Hello", options2);

    worksheet.write_string(CELL("C16"), "This cell comment is scaled in both directions.", text_wrap);

    const xwpp::comment_options_t options3{.x_scale_ = 1.2, .y_scale_ = 0.5};
    worksheet.write_comment(CELL("C16"), "Hello", options3);

    worksheet.write_string(CELL("C19"), "This cell comment has width and height specified in pixels.", text_wrap);

    const xwpp::comment_options_t options4{.width_ = 200, .height_ = 50};
    worksheet.write_comment(CELL("C19"), "Hello", options4);
  }

  // Example 5. Demonstrates changes to the cell comment position.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);
    worksheet.set_row(5, 50);
    worksheet.set_row(8, 50);

    worksheet.show_comments();

    worksheet.write_string(CELL("C3"), "This cell comment is in the default position.", text_wrap);

    worksheet.write_comment(2, 2, "Hello");

    worksheet.write_string(CELL("C6"), "This cell comment has been moved to another cell.", text_wrap);

    const xwpp::comment_options_t options1{.start_row_ = 3, .start_col_ = 4};
    worksheet.write_comment(CELL("C6"), "Hello", options1);

    worksheet.write_string(CELL("C9"), "This cell comment has been shifted within its default cell.", text_wrap);

    const xwpp::comment_options_t options2{.x_offset_ = 30, .y_offset_ = 12};
    worksheet.write_comment(CELL("C9"), "Hello", options2);
  }

  // Example 6. Demonstrates changes to the comment background color.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 50);
    worksheet.set_row(5, 50);
    worksheet.set_row(8, 50);

    worksheet.show_comments();

    worksheet.write_string(CELL("C3"), "This cell comment has a different color.", text_wrap);

    const xwpp::comment_options_t options1{.color_ = xwpp::color_t::GREEN};
    worksheet.write_comment(2, 2, "Hello", options1);

    worksheet.write_string(CELL("C6"), "This cell comment has the default color.", text_wrap);

    worksheet.write_comment(CELL("C6"), "Hello");

    worksheet.write_string(CELL("C9"), "This cell comment has a different color.", text_wrap);

    const xwpp::comment_options_t options2{.color_ = static_cast<xwpp::color_t>(0xFF6600)};
    worksheet.write_comment(CELL("C9"), "Hello", options2);
  }

  // Example 7. Demonstrates how to set the cell comment author.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 30);
    worksheet.set_row(2, 50);
    worksheet.set_row(5, 60);

    worksheet.write_string(CELL("C3"),
                           "Move the mouse over this cell and you will see 'Cell C3 "
                           "commented by' (blank) in the status bar at the bottom.",
                           text_wrap);

    worksheet.write_comment(CELL("C3"), "Hello");

    worksheet.write_string(CELL("C6"),
                           "Move the mouse over this cell and you will see 'Cell C6 "
                           "commented by Xlsxwriter++' in the status bar at the bottom.",
                           text_wrap);

    const xwpp::comment_options_t options{.author_ = "Xlsxwriter++"};
    worksheet.write_comment(CELL("C6"), "Hello", options);
  }

  // Example 8. Demonstrates the need to explicitly set the row height.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.set_column(2, 2, 25);
    worksheet.set_row(2, 80);

    worksheet.show_comments();

    worksheet.write_string(CELL("C3"),
                           "The height of this row has been adjusted explicitly using "
                           "worksheet_set_row(). The size of the comment box is "
                           "adjusted accordingly by Xlsxwriter++",
                           text_wrap);

    worksheet.write_comment(CELL("C3"), "Hello");

    worksheet.write_string(CELL("C6"),
                           "The height of this row has been adjusted by Excel when the "
                           "file is opened due to the text wrap property being set. "
                           "Unfortunately this means that the height of the row is "
                           "unknown to Xlsxwriter++ at run time and thus the comment "
                           "box is stretched as well.\n\n"
                           "Use worksheet_set_row() to specify the row height explicitly "
                           "to avoid this problem.",
                           text_wrap);

    worksheet.write_comment(CELL("C6"), "Hello");
  }

  workbook.save("comments2.xlsx");
}
