/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 */

/**
 * @page format_page The Format object
 *
 * The Format object represents an the formatting properties that can be
 * applied to a cell including: fonts, colors, patterns,
 * borders, alignment and number formatting.
 *
 * See @ref format.h for full details of the functionality.
 *
 * @file
 *
 * @brief Functions and properties for adding formatting to cells in Excel.
 *
 * This section describes the functions and properties that are available for
 * formatting cells in Excel.
 *
 * The properties of a cell that can be formatted include: fonts, colors,
 * patterns, borders, alignment and number formatting.
 *
 * @image html formats_intro.png
 *
 * Formats in `Xlsxwriter++` are accessed via the lxw_format
 * struct. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the workbook_add_format() method as
 * follows:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 * @endcode
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in this section.
 * For example:
 *
 * @code
 *    // Create the Format.
 *    lxw_format *format = workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    format_set_bold(format);
 *    format_set_font_color(format, LXW_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 * The full range of formatting options that can be applied using
 * `Xlsxwriter++` are shown below.
 *
 */

#ifndef XWPP_FORMAT_H
#define XWPP_FORMAT_H

#include <cstdint>
#include <string>

namespace xwpp
{

// TODO Add rgb function to build color_t
/**
 * @brief The type for RGB colors in Xlsxwriter++.
 *
 * The type for RGB colors in Xlsxwriter++. The valid range is `0x000000`
 * (black) to `0xFFFFFF` (white). See @ref working_with_colors.
 *
 * Few predefined colors are provided through enumerates.
 */
enum class color_t : uint32_t
{
  BLACK   = 0x000000,
  BLUE    = 0x0000FF,
  BROWN   = 0x800000,
  CYAN    = 0x00FFFF,
  GRAY    = 0x808080,
  GREEN   = 0x008000,
  LIME    = 0x00FF00,
  MAGENTA = 0xFF00FF,
  NAVY    = 0x000080,
  ORANGE  = 0xFF6600,
  PINK    = 0xFF00FF,
  PURPLE  = 0x800080,
  RED     = 0xFF0000,
  SILVER  = 0xC0C0C0,
  WHITE   = 0xFFFFFF,
  YELLOW  = 0xFFFF00,

  // Special value
  UNSET = 0xFFFFFFFF,
};
const uint32_t COLOR_MASK = 0xFFFFFF;

/// #define LXW_FORMAT_FIELD_LEN            128
const double MIN_FONT_SIZE = 1.0;
const double MAX_FONT_SIZE = 409.0;

// TODO Complete the list
// Format underline values for format_set_underline().
enum class format_underlines_t
{
  NONE = 0,

  // Single underline
  SINGLE,

  // Double underline
  DOUBLE,

  // Single accounting underline

  SINGLE_ACCOUNTING,

  // Double accounting underline
  DOUBLE_ACCOUNTING
};

/** Superscript and subscript values for format_set_font_script(). */
enum class format_scripts_t
{
  /** No script */
  NONE,

  /** Superscript font */
  SUPERSCRIPT,

  /** Subscript font */
  SUBSCRIPT
};

// Alignment values for format_set_align().
enum class format_alignments_t
{
  // No alignment. Cell will use Excel's default for the data type
  NONE = 0,

  // Left horizontal alignment
  HORIZONTAL_LEFT,

  // Center horizontal alignment
  HORIZONTAL_CENTER,

  // Right horizontal alignment
  HORIZONTAL_RIGHT,

  // Cell fill horizontal alignment
  HORIZONTAL_FILL,

  // Justify horizontal alignment
  HORIZONTAL_JUSTIFY,

  // Center Across horizontal alignment
  HORIZONTAL_CENTER_ACROSS,

  // Left horizontal alignment
  HORIZONTAL_DISTRIBUTED,

  // Top vertical alignment
  VERTICAL_TOP,

  // Bottom vertical alignment
  VERTICAL_BOTTOM,

  // Center vertical alignment
  VERTICAL_CENTER,

  // Justify vertical alignment
  VERTICAL_JUSTIFY,

  // Distributed vertical alignment
  VERTICAL_DISTRIBUTED
};

/**
 * Diagonal border types.
 *
 */
enum class format_diagonal_types_t
{
  NONE,

  /** Cell diagonal border from bottom left to top right. */
  BORDER_UP,

  /** Cell diagonal border from top left to bottom right. */
  BORDER_DOWN,

  /** Cell diagonal border in both directions. */
  BORDER_UP_DOWN
};

// Pattern value for use with format_set_pattern().
enum class format_patterns_t
{
  // Empty pattern
  NONE = 0,

  // Solid pattern
  SOLID,

  // Medium gray pattern
  MEDIUM_GRAY,

  // Dark gray pattern
  DARK_GRAY,

  // Light gray pattern
  LIGHT_GRAY,

  // Dark horizontal line pattern
  DARK_HORIZONTAL,

  // Dark vertical line pattern
  DARK_VERTICAL,

  // Dark diagonal stripe pattern
  DARK_DOWN,

  // Reverse dark diagonal stripe pattern
  DARK_UP,

  // Dark grid pattern
  DARK_GRID,

  // Dark trellis pattern
  DARK_TRELLIS,

  // Light horizontal Line pattern
  LIGHT_HORIZONTAL,

  // Light vertical line pattern
  LIGHT_VERTICAL,

  // Light diagonal stripe pattern
  LIGHT_DOWN,

  // Reverse light diagonal stripe pattern
  LIGHT_UP,

  // Light grid pattern
  LIGHT_GRID,

  // Light trellis pattern
  LIGHT_TRELLIS,

  // 12.5% gray pattern
  GRAY_125,

  // 6.25% gray pattern
  GRAY_0625
};

// Cell border styles for use with format_set_border().
enum class format_borders_t
{
  // No border
  NONE,

  // Thin border style
  THIN,

  // Medium border style
  MEDIUM,

  // Dashed border style
  DASHED,

  // Dotted border style
  DOTTED,

  // Thick border style
  THICK,

  // Double border style
  DOUBLE,

  // Hair border style
  HAIR,

  // Medium dashed border style
  MEDIUM_DASHED,

  // Dash-dot border style
  DASH_DOT,

  // Medium dash-dot border style
  MEDIUM_DASH_DOT,

  // Dash-dot-dot border style
  DASH_DOT_DOT,

  // Medium dash-dot-dot border style
  MEDIUM_DASH_DOT_DOT,

  // Slant dash-dot border style
  SLANT_DASH_DOT
};

/**
 * @brief Representation the formatting properties of an Excel format.
 *
 * The members of the format_t struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in format.h.
 *
 * For example:
 *
 * @code
 *    // Create the Format.
 *    format_t *format = workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    format_set_bold(format);
 *    format_set_font_color(format, LXW_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 */
// TODO After creation of higher level class, this should become a struct with all field public
// TODO Or more simplen use std::shared_ptr to exchange with caller
class format_t
{
public:
  /**
   * @brief Set the number format for a cell.
   *
   * @param format      Pointer to a Format instance.
   * @param num_format The cell number format string.
   *
   * This method is used to define the numerical format of a number in
   * Excel. It controls whether a number is displayed as an integer, a
   * floating point number, a date, a currency value or some other user
   * defined format.
   *
   * The numerical format of a cell can be specified by using a format
   * string:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_num_format(format, "d mmm yyyy");
   * @endcode
   *
   * Format strings can control any aspect of number formatting allowed by Excel:
   *
   * @dontinclude format_num_format.c
   * @skipline set_num_format
   * @until 1209
   *
   * @image html format_set_num_format.png
   *
   * To set a number format that matches an Excel format category such as "Date"
   * or "Currency" see @ref ww_formats_categories.
   *
   * The number system used for dates is described in @ref working_with_dates.
   *
   * For more information on number formats in Excel refer to the
   * [Microsoft documentation on cell
   * formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
   */
  void set_num_format(const std::string& num_format);

  /**
   * @brief Turn on bold for the format font.
   *
   * @param format Pointer to a Format instance.
   *
   * Set the bold property of the font:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_bold(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Bold Text", format);
   * @endcode
   *
   * @image html format_font_bold.png
   */
  void set_bold();

  /**
   * @brief Turn on italic for the format font.
   *
   * @param format Pointer to a Format instance.
   *
   * Set the italic property of the font:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_italic(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Italic Text", format);
   * @endcode
   *
   * @image html format_font_italic.png
   */
  void set_italic();

  /**
   * @brief Set the alignment for data in the cell.
   *
   * @param format    Pointer to a Format instance.
   * @param alignment The horizontal and or vertical alignment direction.
   *
   * This method is used to set the horizontal and vertical text alignment within
   * a cell. The following are the available horizontal alignments:
   *
   * - #LXW_ALIGN_LEFT
   * - #LXW_ALIGN_CENTER
   * - #LXW_ALIGN_RIGHT
   * - #LXW_ALIGN_FILL
   * - #LXW_ALIGN_JUSTIFY
   * - #LXW_ALIGN_CENTER_ACROSS
   * - #LXW_ALIGN_DISTRIBUTED
   *
   * The following are the available vertical alignments:
   *
   * - #LXW_ALIGN_VERTICAL_TOP
   * - #LXW_ALIGN_VERTICAL_BOTTOM
   * - #LXW_ALIGN_VERTICAL_CENTER
   * - #LXW_ALIGN_VERTICAL_JUSTIFY
   * - #LXW_ALIGN_VERTICAL_DISTRIBUTED
   *
   * As in Excel, vertical and horizontal alignments can be combined:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *
   *     format_set_align(format, LXW_ALIGN_CENTER);
   *     format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
   *
   *     worksheet_set_row(0, 30);
   *     worksheet_write_string(worksheet, 0, 0, "Some Text", format);
   * @endcode
   *
   * @image html format_font_align.png
   *
   * Text can be aligned across two or more adjacent cells using the
   * center_across property. However, for genuine merged cells it is better to
   * use the worksheet_merge_range() worksheet method.
   *
   * The vertical justify option can be used to provide automatic text wrapping
   * in a cell. The height of the cell will be adjusted to accommodate the
   * wrapped text. To specify where the text wraps use the
   * format_set_text_wrap() method.
   */
  void set_align(format_alignments_t alignment);

  /**
   * @brief Set the color of the font used in the cell.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell font color.
   *
   *
   * Set the font color:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_font_color(format, LXW_COLOR_RED);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Wheelbarrow", format);
   * @endcode
   *
   * @image html format_font_color.png
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   *
   * @note
   * The format_set_font_color() method is used to set the font color in a
   * cell. To set the color of a cell background use the format_set_bg_color()
   * and format_set_pattern() methods.
   */
  void set_font_color(color_t color);

  /**
   * @brief Set the pattern foreground color for a cell.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell pattern foreground  color.
   *
   * The format_set_fg_color() method can be used to set the foreground color of
   * a pattern.
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   *
   */
  void set_fg_color(color_t color);

  /**
   * @brief Turn on underline for the format:
   *
   * @param format Pointer to a Format instance.
   * @param style Underline style.
   *
   * Set the underline property of the format:
   *
   * @code
   *     format_set_underline(format, LXW_UNDERLINE_SINGLE);
   * @endcode
   *
   * @image html format_font_underlined.png
   *
   * The available underline styles are:
   *
   * - #LXW_UNDERLINE_SINGLE
   * - #LXW_UNDERLINE_DOUBLE
   * - #LXW_UNDERLINE_SINGLE_ACCOUNTING
   * - #LXW_UNDERLINE_DOUBLE_ACCOUNTING
   *
   */
  void set_underline(format_underlines_t style);

  /**
   * @brief Wrap text in a cell.
   *
   * Turn text wrapping on for text in a cell.
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_text_wrap(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a
   * cell", format);
   * @endcode
   *
   * If you wish to control where the text is wrapped you can add newline
   * characters to the string:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_text_wrap(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format);
   * @endcode
   *
   * @image html format_font_text_wrap.png
   *
   * Excel will adjust the height of the row to accommodate the wrapped text. A
   * similar effect can be obtained without newlines using the
   * format_set_align() function with #LXW_ALIGN_VERTICAL_JUSTIFY.
   */
  void set_text_wrap();

  void set_hyperlink();

  /**
   * @brief Set the cell border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  Border style index.
   *
   * Set the cell border style:
   *
   * @code
   *     format_set_border(format, LXW_BORDER_THIN);
   * @endcode
   *
   * Individual border elements can be configured using the following functions
   * with the same parameters:
   *
   * - format_set_bottom()
   * - format_set_top()
   * - format_set_left()
   * - format_set_right()
   *
   * A cell border is comprised of a border on the bottom, top, left and right.
   * These can be set to the same value using format_set_border() or
   * individually using the relevant method calls shown above.
   *
   * The following border styles are available:
   *
   * - #LXW_BORDER_THIN
   * - #LXW_BORDER_MEDIUM
   * - #LXW_BORDER_DASHED
   * - #LXW_BORDER_DOTTED
   * - #LXW_BORDER_THICK
   * - #LXW_BORDER_DOUBLE
   * - #LXW_BORDER_HAIR
   * - #LXW_BORDER_MEDIUM_DASHED
   * - #LXW_BORDER_DASH_DOT
   * - #LXW_BORDER_MEDIUM_DASH_DOT
   * - #LXW_BORDER_DASH_DOT_DOT
   * - #LXW_BORDER_MEDIUM_DASH_DOT_DOT
   * - #LXW_BORDER_SLANT_DASH_DOT
   *
   *  The most commonly used style is the `thin` style.
   */
  void set_border(format_borders_t style);

  /**
   * @brief Set the cell bottom border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  Border style index.
   *
   * Set the cell bottom border style. See format_set_border() for details on the
   * border styles.
   */
  void set_bottom(format_borders_t style);

  /**
   * @brief Set the cell top border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  Border style index.
   *
   * Set the cell top border style. See format_set_border() for details on the
   * border styles.
   */
  void set_top(format_borders_t style);

  /**
   * @brief Set the cell left border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  Border style index.
   *
   * Set the cell left border style. See format_set_border() for details on the
   * border styles.
   */
  void set_left(format_borders_t style);

  /**
   * @brief Set the cell right border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  Border style index.
   *
   * Set the cell right border style. See format_set_border() for details on the
   * border styles.
   */
  void set_right(format_borders_t style);

  /**
   * @brief Set the pattern background color for a cell.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell pattern background color.
   *
   * The format_set_bg_color() method can be used to set the background color of
   * a pattern. Patterns are defined via the format_set_pattern() method. If a
   * pattern hasn't been defined then a solid fill pattern is used as the
   * default.
   *
   * Here is an example of how to set up a solid fill in a cell:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *
   *     format_set_pattern (format, LXW_PATTERN_SOLID);
   *     format_set_bg_color(format, LXW_COLOR_GREEN);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Ray", format);
   * @endcode
   *
   * @image html formats_set_bg_color.png
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   *
   */
  void set_bg_color(color_t color);

  /**
   * @brief Set the background fill pattern for a cell
   *
   * @param format Pointer to a Format instance.
   * @param index  Pattern index.
   *
   * Set the background pattern for a cell.
   *
   * The most common pattern is a solid fill of the background color:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *
   *     format_set_pattern (format, LXW_PATTERN_SOLID);
   *     format_set_bg_color(format, LXW_COLOR_YELLOW);
   * @endcode
   *
   * The available fill patterns are:
   *
   *    Fill Type                     | Define
   *    ----------------------------- | -----------------------------
   *    Solid                         | #LXW_PATTERN_SOLID
   *    Medium gray                   | #LXW_PATTERN_MEDIUM_GRAY
   *    Dark gray                     | #LXW_PATTERN_DARK_GRAY
   *    Light gray                    | #LXW_PATTERN_LIGHT_GRAY
   *    Dark horizontal line          | #LXW_PATTERN_DARK_HORIZONTAL
   *    Dark vertical line            | #LXW_PATTERN_DARK_VERTICAL
   *    Dark diagonal stripe          | #LXW_PATTERN_DARK_DOWN
   *    Reverse dark diagonal stripe  | #LXW_PATTERN_DARK_UP
   *    Dark grid                     | #LXW_PATTERN_DARK_GRID
   *    Dark trellis                  | #LXW_PATTERN_DARK_TRELLIS
   *    Light horizontal line         | #LXW_PATTERN_LIGHT_HORIZONTAL
   *    Light vertical line           | #LXW_PATTERN_LIGHT_VERTICAL
   *    Light diagonal stripe         | #LXW_PATTERN_LIGHT_DOWN
   *    Reverse light diagonal stripe | #LXW_PATTERN_LIGHT_UP
   *    Light grid                    | #LXW_PATTERN_LIGHT_GRID
   *    Light trellis                 | #LXW_PATTERN_LIGHT_TRELLIS
   *    12.5% gray                    | #LXW_PATTERN_GRAY_125
   *    6.25% gray                    | #LXW_PATTERN_GRAY_0625
   *
   */
  void set_pattern(format_patterns_t pattern);

  /**
   * @brief Set the color of the cell border.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell border color.
   *
   * Individual border elements can be configured using the following methods with
   * the same parameters:
   *
   * - format_set_bottom_color()
   * - format_set_top_color()
   * - format_set_left_color()
   * - format_set_right_color()
   *
   * Set the color of the cell borders. A cell border is comprised of a border
   * on the bottom, top, left and right. These can be set to the same color
   * using format_set_border_color() or individually using the relevant method
   * calls shown above.
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_border_color(color_t color);

  /**
   * @brief Set the color of the bottom cell border.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell border color.
   *
   * See format_set_border_color() for details on the border colors.
   */
  void set_bottom_color(color_t color);

  /**
   * @brief Set the color of the top cell border.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell border color.
   *
   * See format_set_border_color() for details on the border colors.
   */
  void set_top_color(color_t color);

  /**
   * @brief Set the color of the left cell border.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell border color.
   *
   * See format_set_border_color() for details on the border colors.
   */
  void set_left_color(color_t color);

  /**
   * @brief Set the color of the right cell border.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell border color.
   *
   * See format_set_border_color() for details on the border colors.
   */
  void set_right_color(color_t color);

  /**
   * @brief Set the diagonal cell border type.
   *
   * @param format Pointer to a Format instance.
   * @param type   The #lxw_format_diagonal_types diagonal border type.
   *
   * Set the diagonal cell border type:
   *
   * @code
   *     lxw_format *format1 = workbook_add_format(workbook);
   *     format_set_diag_type(  format1, LXW_DIAGONAL_BORDER_UP);
   *
   *     lxw_format *format2 = workbook_add_format(workbook);
   *     format_set_diag_type(  format2, LXW_DIAGONAL_BORDER_DOWN);
   *
   *     lxw_format *format3 = workbook_add_format(workbook);
   *     format_set_diag_type(  format3, LXW_DIAGONAL_BORDER_UP_DOWN);
   *
   *     lxw_format *format4 = workbook_add_format(workbook);
   *     format_set_diag_type(  format4, LXW_DIAGONAL_BORDER_UP_DOWN);
   *     format_set_diag_border(format4, LXW_BORDER_HAIR);
   *     format_set_diag_color( format4, LXW_COLOR_RED);
   *
   *     worksheet_write_string(worksheet, CELL("B3"),  "Text", format1);
   *     worksheet_write_string(worksheet, CELL("B6"),  "Text", format2);
   *     worksheet_write_string(worksheet, CELL("B9"),  "Text", format3);
   *     worksheet_write_string(worksheet, CELL("B12"), "Text", format4);
   * @endcode
   *
   * @image html diagonal_border.png
   *
   * The allowable border types are defined in #lxw_format_diagonal_types:
   *
   * - #LXW_DIAGONAL_BORDER_UP: Cell diagonal border from bottom left to top
   *   right.
   *
   * - #LXW_DIAGONAL_BORDER_DOWN: Cell diagonal border from top left to bottom
   *   right.
   *
   * - #LXW_DIAGONAL_BORDER_UP_DOWN: Cell diagonal border from top left to
   *   bottom right. A combination of the 2 previous types.
   *
   * If the border style isn't specified with `format_set_diag_border()` then it
   * will default to #LXW_BORDER_THIN.
   */
  void set_diag_type(format_diagonal_types_t type);

  /**
   * @brief Set the diagonal cell border style.
   *
   * @param format Pointer to a Format instance.
   * @param style  The #lxw_format_borders style.
   *
   * Set the diagonal border style. This should be a #lxw_format_borders value.
   * See the example above.
   *
   */
  void set_diag_border(format_borders_t style);

  /**
   * @brief Set the diagonal cell border color.
   *
   * @param format Pointer to a Format instance.
   * @param color  The cell diagonal border color.
   *
   * Set the diagonal border color. The color should be an RGB integer value,
   * see @ref working_with_colors and the above example.
   */
  void set_diag_color(color_t color);

  /**
   * @brief Set the Excel built-in number format for a cell.
   *
   * @param format Pointer to a Format instance.
   * @param index  The built-in number format index for the cell.
   *
   * This function is similar to format_set_num_format() except that it takes an
   * index to a limited number of Excel's built-in number formats instead of a
   * user defined format string:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_num_format_index(format, 0x0F); // d-mmm-yy
   * @endcode
   *
   * @note
   * Unless you need to specifically access one of Excel's built-in number
   * formats the format_set_num_format() function above is a better
   * solution. The format_set_num_format_index() function is mainly included for
   * backward compatibility and completeness.
   *
   * The Excel built-in number formats as shown in the table below:
   *
   *   | Index | Index | Format String                                        |
   *   | ----- | ----- | ---------------------------------------------------- |
   *   | 0     | 0x00  | `General`                                            |
   *   | 1     | 0x01  | `0`                                                  |
   *   | 2     | 0x02  | `0.00`                                               |
   *   | 3     | 0x03  | `#,##0`                                              |
   *   | 4     | 0x04  | `#,##0.00`                                           |
   *   | 5     | 0x05  | `($#,##0_);($#,##0)`                                 |
   *   | 6     | 0x06  | `($#,##0_);[Red]($#,##0)`                            |
   *   | 7     | 0x07  | `($#,##0.00_);($#,##0.00)`                           |
   *   | 8     | 0x08  | `($#,##0.00_);[Red]($#,##0.00)`                      |
   *   | 9     | 0x09  | `0%`                                                 |
   *   | 10    | 0x0a  | `0.00%`                                              |
   *   | 11    | 0x0b  | `0.00E+00`                                           |
   *   | 12    | 0x0c  | `# ?/?`                                              |
   *   | 13    | 0x0d  | `# ??/??`                                            |
   *   | 14    | 0x0e  | `m/d/yy`                                             |
   *   | 15    | 0x0f  | `d-mmm-yy`                                           |
   *   | 16    | 0x10  | `d-mmm`                                              |
   *   | 17    | 0x11  | `mmm-yy`                                             |
   *   | 18    | 0x12  | `h:mm AM/PM`                                         |
   *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
   *   | 20    | 0x14  | `h:mm`                                               |
   *   | 21    | 0x15  | `h:mm:ss`                                            |
   *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
   *   | ...   | ...   | ...                                                  |
   *   | 37    | 0x25  | `(#,##0_);(#,##0)`                                   |
   *   | 38    | 0x26  | `(#,##0_);[Red](#,##0)`                              |
   *   | 39    | 0x27  | `(#,##0.00_);(#,##0.00)`                             |
   *   | 40    | 0x28  | `(#,##0.00_);[Red](#,##0.00)`                        |
   *   | 41    | 0x29  | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
   *   | 42    | 0x2a  | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
   *   | 43    | 0x2b  | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
   *   | 44    | 0x2c  | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
   *   | 45    | 0x2d  | `mm:ss`                                              |
   *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
   *   | 47    | 0x2f  | `mm:ss.0`                                            |
   *   | 48    | 0x30  | `##0.0E+0`                                           |
   *   | 49    | 0x31  | `@`                                                  |
   *
   * @note
   *  -  Numeric formats 23 to 36 are not documented by Microsoft and may differ
   *     in international versions. The listed date and currency formats may also
   *     vary depending on system settings.
   *  - The dollar sign in the above format appears as the defined local currency
   *    symbol.
   *  - These formats can also be set via format_set_num_format().
   *  - See also @ref ww_formats_categories.
   */
  // TODO Use dedicated type ?
  void set_num_format_index(uint8_t index);

  /**
   * @brief Set the font used in the cell.
   *
   * @param format    Pointer to a Format instance.
   * @param font_name Cell font name.
   *
   * Specify the font used used in the cell format:
   *
   * @code
   *     format_set_font_name(format, "Avenir Black Oblique");
   * @endcode
   *
   * @image html format_set_font_name.png
   *
   * Excel can only display fonts that are installed on the system that it is
   * running on. Therefore it is generally best to use the fonts that come as
   * standard with Excel such as Calibri, Times New Roman and Courier New.
   *
   * The default font in Excel 2007, and later, is Calibri.
   */
  void set_font_name(const std::string& font_name);

  /**
   * @brief Set the size of the font used in the cell.
   *
   * @param format Pointer to a Format instance.
   * @param size   The cell font size.
   *
   * Set the font size of the cell format:
   *
   * @code
   *     format_set_font_size(format, 30);
   * @endcode
   *
   * @image html format_font_size.png
   *
   * Excel adjusts the height of a row to accommodate the largest font
   * size in the row. You can also explicitly specify the height of a
   * row using the worksheet_set_row() function.
   */
  void set_font_size(double size);

  /**
   * @brief Set the strikeout property of the font.
   *
   * @param format Pointer to a Format instance.
   *
   * @image html format_font_strikeout.png
   *
   */
  void set_font_strikeout();

  /**
   * @brief Set the superscript/subscript property of the font.
   *
   * @param format Pointer to a Format instance.
   * @param style  Superscript or subscript style.
   *
   * Set the superscript o subscript property of the font.
   *
   * @image html format_font_script.png
   *
   * The available script styles are:
   *
   * - #LXW_FONT_SUPERSCRIPT
   * - #LXW_FONT_SUBSCRIPT
   */
  void set_font_script(format_scripts_t style);

  /**
   * @brief Set the Format font family property.
   *
   * @param format Pointer to a Format instance.
   * @param value  The font family index.
   *
   * Set the font family. This is usually an integer in the range 1-4. This
   * function is implemented for completeness but is rarely used in practice.
   *
   * @code
   *     format_set_font_family(format, 178);
   * @endcode
   *
   */
  void set_font_family(uint8_t value);

  /**
   * @brief Set the Format font character set property.
   *
   * @param format Pointer to a Format instance.
   * @param value  The font character set.
   *
   * Set the font character set property. This function is implemented for
   * completeness but is rarely used in practice.
   *
   * @code
   *     format_set_font_charset(format, 178);
   * @endcode
   *
   */
  void set_font_charset(uint8_t value);

  void set_font_outline();
  void set_font_shadow();
  void set_font_scheme(const std::string& font_scheme);
  void set_font_condense();
  void set_font_extend();
  void set_font_only();

  /**
   * @brief Set the cell unlocked state.
   *
   * @param format Pointer to a Format instance.
   *
   * This property can be used to allow modification of a cell in a protected
   * worksheet. In Excel, cell locking is turned on by default for all
   * cells. However, it only has an effect if the worksheet has been protected
   * using the worksheet worksheet_protect() function:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_unlocked(format);
   *
   *     // Enable worksheet protection, without password or options.
   *     worksheet_protect(worksheet, NULL, NULL);
   *
   *     // This cell cannot be edited.
   *     worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL);
   *
   *     // This cell can be edited.
   *     worksheet_write_formula(worksheet, 1, 0, "=1+2", format);
   * @endcode
   */
  void set_unlocked();

  /**
   * @brief Hide formulas in a cell.
   *
   * @param format Pointer to a Format instance.
   *
   * This property is used to hide a formula while still displaying its
   * result. This is generally used to hide complex calculations from end users
   * who are only interested in the result. It only has an effect if the
   * worksheet has been protected using the worksheet worksheet_protect()
   * function:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_hidden(format);
   *
   *     // Enable worksheet protection, without password or options.
   *     worksheet_protect(worksheet, NULL, NULL);
   *
   *     // The formula in this cell isn't visible.
   *     worksheet_write_formula(worksheet, 0, 0, "=1+2", format);
   * @endcode
   */
  void set_hidden();

  /**
   * @brief Set the rotation of the text in a cell.
   *
   * @param format Pointer to a Format instance.
   * @param angle  Rotation angle in the range -90 to 90 and 270.
   *
   * Set the rotation of the text in a cell. The rotation can be any angle in the
   * range -90 to 90 degrees:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_rotation(format, 30);
   *
   *     worksheet_write_string(worksheet, 0, 0, "This text is rotated", format);
   * @endcode
   *
   * @image html format_font_text_rotated.png
   *
   * The angle 270 is also supported. This indicates text where the letters run
   * from top to bottom.
   */
  void set_rotation(int16_t angle);

  /**
   * @brief Set the cell text indentation level.
   *
   * @param format Pointer to a Format instance.
   * @param level  Indentation level.
   *
   * This method can be used to indent text in a cell. The argument, which should
   * be an integer, is taken as the level of indentation:
   *
   * @code
   *     format1 = workbook_add_format(workbook);
   *     format2 = workbook_add_format(workbook);
   *
   *     format_set_indent(format1, 1);
   *     format_set_indent(format2, 2);
   *
   *     worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",
   * format1); worksheet_write_string(worksheet, 1, 0, "This text is indented 2
   * levels", format2);
   * @endcode
   *
   * @image html text_indent.png
   *
   * @note
   * Indentation is a horizontal alignment property. It will override any other
   * horizontal properties but it can be used in conjunction with vertical
   * properties.
   */
  void set_indent(uint8_t level);

  /**
   * @brief Turn on the text "shrink to fit" for a cell.
   *
   * @param format Pointer to a Format instance.
   *
   * This method can be used to shrink text so that it fits in a cell:
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_shrink(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!",
   * format);
   * @endcode
   */
  void set_shrink();

  void set_text_justlast();
  void set_valign(format_alignments_t alignment);
  void set_halign(format_alignments_t alignment);
  void set_reading_order(uint8_t value);
  void set_theme(uint8_t value);
  void set_color_indexed(uint8_t value);

  /**
   * @brief Turn on quote prefix for the format.
   *
   * @param format Pointer to a Format instance.
   *
   * Set the quote prefix property of a format to ensure a string is treated
   * as a string after editing. This is the same as prefixing the string with
   * a single quote in Excel. You don't need to add the quote to the
   * string but you do need to add the format.
   *
   * @code
   *     format = workbook_add_format(workbook);
   *     format_set_quote_prefix(format);
   *
   *     worksheet_write_string(worksheet, 0, 0, "=Foo", format);
   * @endcode
   *
   */
  void set_quote_prefix();

  static const int32_t PROPERTY_UNSET = -1;
  static const std::string DEFAULT_FONT_NAME;

private:
  // TODO friend up to refactoring with struct / class
  friend class workbook_t;
  friend class worksheet_t; // TODO for debug
  friend class style_t;

  [[nodiscard]] format_borders_t check_border(format_borders_t style) const;

  static const uint8_t DEFAULT_FONT_FAMILY = 2;

  ///  lxw_hash_table *xf_format_indices_ = nullptr;
  ///  lxw_hash_table *dxf_format_indices_ = nullptr;
  ///  uint16_t *num_xf_formats_;
  ///  uint16_t *num_dxf_formats_;
  int32_t xf_index_  = PROPERTY_UNSET;
  int32_t dxf_index_ = PROPERTY_UNSET;
  int32_t xf_id_     = 0;
  std::string num_format_;
  std::string font_name_;
  std::string font_scheme_;
  int32_t num_format_index_          = 0;
  int32_t font_index_                = PROPERTY_UNSET;
  bool has_font_                     = false;
  bool has_dxf_font_                 = false;
  double font_size_                  = 11.0;
  bool bold_                         = false;
  bool italic_                       = false;
  color_t font_color_                = color_t::UNSET;
  format_underlines_t underline_     = format_underlines_t::NONE;
  bool font_strikeout_               = false;
  bool font_outline_                 = false;
  bool font_shadow_                  = false;
  format_scripts_t font_script_      = format_scripts_t::NONE;
  uint8_t font_family_               = DEFAULT_FONT_FAMILY;
  bool font_charset_                 = false;
  bool font_condense_                = false;
  bool font_extend_                  = false;
  uint8_t theme_                     = 0;
  bool hyperlink_                    = false;
  bool hidden_                       = false;
  bool locked_                       = true;
  format_alignments_t text_h_align_  = format_alignments_t::NONE;
  bool text_wrap_                    = false;
  format_alignments_t text_v_align_  = format_alignments_t::NONE;
  bool text_justlast_                = false;
  int16_t rotation_                  = 0;
  color_t fg_color_                  = color_t::UNSET;
  color_t bg_color_                  = color_t::UNSET;
  color_t dxf_fg_color_              = color_t::UNSET;
  color_t dxf_bg_color_              = color_t::UNSET;
  format_patterns_t pattern_         = format_patterns_t::NONE;
  bool has_fill_                     = false;
  bool has_dxf_fill_                 = false;
  int32_t fill_index_                = PROPERTY_UNSET;
  int32_t fill_count_                = 0;
  int32_t border_index_              = PROPERTY_UNSET;
  bool has_border_                   = false;
  bool has_dxf_border_               = false;
  int32_t border_count_              = 0;
  format_borders_t bottom_           = format_borders_t::NONE;
  format_borders_t diag_border_      = format_borders_t::NONE;
  format_diagonal_types_t diag_type_ = format_diagonal_types_t::NONE;
  format_borders_t left_             = format_borders_t::NONE;
  format_borders_t right_            = format_borders_t::NONE;
  format_borders_t top_              = format_borders_t::NONE;
  color_t bottom_color_              = color_t::UNSET;
  color_t diag_color_                = color_t::UNSET;
  color_t left_color_                = color_t::UNSET;
  color_t right_color_               = color_t::UNSET;
  color_t top_color_                 = color_t::UNSET;
  uint8_t indent_                    = 0;
  bool shrink_                       = false;
  bool merge_range_                  = false;
  uint8_t reading_order_             = 0;
  bool just_distrib_                 = false;
  uint8_t color_indexed_             = 0;
  bool font_only_                    = false;
  bool quote_prefix_                 = false;
};

/*
 * Struct to represent the font component of a format.
 */
/// typedef struct lxw_font {

///  char font_name[LXW_FORMAT_FIELD_LEN];
///  double font_size;
///  uint8_t bold;
///  uint8_t italic;
///  uint8_t underline;
///  uint8_t theme;
///  uint8_t font_strikeout;
///  uint8_t font_outline;
///  uint8_t font_shadow;
///  uint8_t font_script;
///  uint8_t font_family;
///  uint8_t font_charset;
///  uint8_t font_condense;
///  uint8_t font_extend;
///  lxw_color_t font_color;
/// } lxw_font;

/*
 * Struct to represent the border component of a format.
 */
/// typedef struct lxw_border {
///  uint8_t bottom;
///  uint8_t diag_border;
///  uint8_t diag_type;
///  uint8_t left;
///  uint8_t right;
///  uint8_t top;

///  lxw_color_t bottom_color;
///  lxw_color_t diag_color;
///  lxw_color_t left_color;
///  lxw_color_t right_color;
///  lxw_color_t top_color;

/// } lxw_border;

/*
 * Struct to represent the fill component of a format.
 */
/// typedef struct lxw_fill {

///  lxw_color_t fg_color;
///  lxw_color_t bg_color;
///  uint8_t pattern;

/// } lxw_fill;

/// int32_t lxw_format_get_dxf_index(lxw_format *format);
/// lxw_border *lxw_format_get_border_key(lxw_format *format);
/// lxw_fill *lxw_format_get_fill_key(lxw_format *format);

}

#endif
