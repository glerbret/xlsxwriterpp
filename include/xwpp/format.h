/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page format_page The Format object
 *
 * The Format object represents an the formatting properties that can be
 * applied to a cell including: fonts, colors, patterns,
 * borders, alignment and number formatting.
 *
 * @see @ref format.h for full details of the functionality.
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
 * Formats in `Xlsxwriter++` are accessed via the `format_t`
 * class. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the `workbook_t::add_format()` method as
 * follows:
 *
 * @code
 *  // Create the Format.
 *  xwpp::format_t* format = workbook.add_format();
 *
 *  // Set some of the format properties.
 *  format->set_bold();
 *  format->set_font_color(xwpp::color_t::RED);
 *
 *  // Use the format to change the text format in a cell.
 *  worksheet.write_string(0, 0, "Hello", format);
 * @endcode
 */

#ifndef XWPP_FORMAT_H
#define XWPP_FORMAT_H

#include <cstdint>
#include <functional>
#include <string>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief The type for RGB colors in `Xlsxwriter++`.
 *
 * The type for RGB colors in `Xlsxwriter++`. The valid range is `0x000000`
 * (black) to `0xFFFFFF` (white).
 *
 * @see @ref working_with_colors.
 *
 * Few predefined colors are provided through enumerates.
 *
 * @todo Replace by a class or struct.
 * @todo With predefined value in an enum class, function to build `%color_t`, ...
 *
 * @showenumvalues
 */
enum class color_t : uint32_t
{
  /** <span style="color:#000000">Black.</span>*/
  BLACK   = 0x000000,
  /** <span style="color:#0000FF">Blue.</span>*/
  BLUE    = 0x0000FF,
  /** <span style="color:#800000">Brown.</span>*/
  BROWN   = 0x800000,
  /** <span style="color:#00FFFF">Cyan.</span>*/
  CYAN    = 0x00FFFF,
  /** <span style="color:#808080">Gray.</span>*/
  GRAY    = 0x808080,
  /** <span style="color:#008000">Green.</span>*/
  GREEN   = 0x008000,
  /** <span style="color:#00FF00">Lime.</span>*/
  LIME    = 0x00FF00,
  /** <span style="color:#FF00FF">Magenta.</span>*/
  MAGENTA = 0xFF00FF,
  /** <span style="color:#000080">Navy.</span>*/
  NAVY    = 0x000080,
  /** <span style="color:#FF6600">Orange.</span>*/
  ORANGE  = 0xFF6600,
  /** <span style="color:#FF00FF">Pink.</span>*/
  PINK    = 0xFF00FF,
  /** <span style="color:#800080">Purple.</span>*/
  PURPLE  = 0x800080,
  /** <span style="color:#FF0000">Red.</span>*/
  RED     = 0xFF0000,
  /** <span style="color:#C0C0C0">Silver.</span>*/
  SILVER  = 0xC0C0C0,
  /** <span style="color:#FFFFFF">White.</span>*/
  WHITE   = 0xFFFFFF,
  /** <span style="color:#FFFF00">Yellow.</span>*/
  YELLOW  = 0xFFFF00,

  // Special value
  UNSET = 0xFFFFFFFF,
};
const uint32_t COLOR_MASK = 0xFFFFFF;

const double MIN_FONT_SIZE = 1.0;
const double MAX_FONT_SIZE = 409.0;

/**
 * @brief Option for underline format.
 *
 * Format underline values for `format_t::set_underline()`.
 *
 * @todo Complete the list
 */
enum class format_underlines_t
{
  NONE = 0,

  /** Single underline. */
  SINGLE,

  /** Double underline. */
  DOUBLE,

  /** Single accounting underline. */
  SINGLE_ACCOUNTING,

  /** Double accounting underline. */
  DOUBLE_ACCOUNTING
};

/**
 * @brief Option for superscript and subscript.
 *
 * Superscript and subscript values for `format_t::set_font_script()`.
 */
enum class format_scripts_t
{
  /** No script. */
  NONE,

  /** Superscript font. */
  SUPERSCRIPT,

  /** Subscript font. */
  SUBSCRIPT
};

/**
 * @brief Alignment options.
 *
 * Alignment values for `format_t::set_align()`.
 *
 * @todo Separate vertical / horizontal alignment option?
 */
enum class format_alignments_t
{
  /** No alignment. Cell will use Excel's default for the data type. */
  NONE = 0,

  /** Left horizontal alignment. */
  HORIZONTAL_LEFT,

  /** Center horizontal alignment. */
  HORIZONTAL_CENTER,

  /** Right horizontal alignment. */
  HORIZONTAL_RIGHT,

  /** Cell fill horizontal alignment. */
  HORIZONTAL_FILL,

  /** Justify horizontal alignment. */
  HORIZONTAL_JUSTIFY,

  /** Center Across horizontal alignment. */
  HORIZONTAL_CENTER_ACROSS,

  /** Left horizontal alignment. */
  HORIZONTAL_DISTRIBUTED,

  /** Top vertical alignment. */
  VERTICAL_TOP,

  /** Bottom vertical alignment. */
  VERTICAL_BOTTOM,

  /** Center vertical alignment. */
  VERTICAL_CENTER,

  /** Justify vertical alignment. */
  VERTICAL_JUSTIFY,

  /** Distributed vertical alignment. */
  VERTICAL_DISTRIBUTED
};

/**
 * @brief Diagonal border types.
 *
 * Diagonal border types.
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

/**
 * @brief Pattern options.
 *
 * Pattern value for use with `format_t::set_pattern()`.
 */
enum class format_patterns_t
{
  /** Empty pattern. */
  NONE = 0,

  /** Solid pattern. */
  SOLID,

  /** Medium gray pattern. */
  MEDIUM_GRAY,

  /** Dark gray pattern. */
  DARK_GRAY,

  /** Light gray pattern. */
  LIGHT_GRAY,

  /** Dark horizontal line pattern. */
  DARK_HORIZONTAL,

  /** Dark vertical line pattern. */
  DARK_VERTICAL,

  /** Dark diagonal stripe pattern. */
  DARK_DOWN,

  /** Reverse dark diagonal stripe pattern. */
  DARK_UP,

  /** Dark grid pattern. */
  DARK_GRID,

  /** Dark trellis pattern. */
  DARK_TRELLIS,

  /** Light horizontal Line pattern. */
  LIGHT_HORIZONTAL,

  /** Light vertical line pattern. */
  LIGHT_VERTICAL,

  /** Light diagonal stripe pattern. */
  LIGHT_DOWN,

  /** Reverse light diagonal stripe pattern. */
  LIGHT_UP,

  /** Light grid pattern. */
  LIGHT_GRID,

  /** Light trellis pattern. */
  LIGHT_TRELLIS,

  /** 12.5% gray pattern. */
  GRAY_125,

  /** 6.25% gray pattern. */
  GRAY_0625
};

/**
 * @brief Cell border styles.
 *
 * Cell border styles for use with `format_t::set_border()`.
 */
enum class format_borders_t
{
  /** No border. */
  NONE,

  /** Thin border style. */
  THIN,

  /** Medium border style. */
  MEDIUM,

  /** Dashed border style. */
  DASHED,

  /** Dotted border style. */
  DOTTED,

  /** Thick border style. */
  THICK,

  /** Double border style. */
  DOUBLE,

  /** Hair border style. */
  HAIR,

  /** Medium dashed border style. */
  MEDIUM_DASHED,

  /** Dash-dot border style. */
  DASH_DOT,

  /** Medium dash-dot border style. */
  MEDIUM_DASH_DOT,

  /** Dash-dot-dot border style. */
  DASH_DOT_DOT,

  /** Medium dash-dot-dot border style. */
  MEDIUM_DASH_DOT_DOT,

  /** Slant dash-dot border style. */
  SLANT_DASH_DOT
};

/**
 * @brief Representation the formatting properties of an Excel format.
 *
 * Representation the formatting properties of an Excel format with API to
 * change it.
 *
 * @code
 *  // Create the Format.
 *  xwpp::format_t* format = workbook.add_format();
 *
 *  // Set some of the format properties.
 *  format->set_bold();
 *  format->set_font_color(xwpp::color_t::RED);
 *
 *  // Use the format to change the text format in a cell.
 *  worksheet.write_string(0, 0, "Hello", format);
 * @endcode
 *
 * @todo Rework this to improve format (builder, encapsulation in higher-level class,
 * independent class, manage it through std::shared_ptr, ...).
 * @todo Add builder API.
 */
class format_t
{
public:
  explicit format_t(std::function<int32_t(format_t*)> get_dxf_index);

  /**
   * @brief Set the number format for a cell.
   *
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
   *  xwpp::format_t* format = workbook.add_format();
   *  format->set_num_format("d mmm yyyy");
   * @endcode
   *
   * Format strings can control any aspect of number formatting allowed by Excel:
   *
   * @dontinclude format_num_format.cpp
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
   * Set the bold property of the font:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_bold();
   *
   *  worksheet.write_string(0, 0, "Bold Text", format);
   * @endcode
   *
   * @image html format_font_bold.png
   */
  void set_bold();

  /**
   * @brief Turn on italic for the format font.
   *
   * Set the italic property of the font:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_italic();
   *
   *  worksheet.write_string(0, 0, "Italic Text", format);
   * @endcode
   *
   * @image html format_font_italic.png
   */
  void set_italic();

  /**
   * @brief Set the alignment for data in the cell.
   *
   * @param alignment The horizontal and / or vertical alignment direction.
   *
   * This method is used to set the horizontal and vertical text alignment within
   * a cell. The following are the available horizontal alignments:
   *
   * - `%format_alignments_t::HORIZONTAL_LEFT`
   * - `%format_alignments_t::HORIZONTAL_CENTER`
   * - `%format_alignments_t::HORIZONTAL_RIGHT`
   * - `%format_alignments_t::HORIZONTAL_FILL`
   * - `%format_alignments_t::HORIZONTAL_JUSTIFY`
   * - `%format_alignments_t::HORIZONTAL_CENTER_ACROSS`
   * - `%format_alignments_t::HORIZONTAL_DISTRIBUTED`
   *
   * The following are the available vertical alignments:
   *
   * - `%format_alignments_t::VERTICAL_TOP`
   * - `%format_alignments_t::VERTICAL_BOTTOM`
   * - `%format_alignments_t::VERTICAL_CENTER`
   * - `%format_alignments_t::VERTICAL_JUSTIFY`
   * - `%format_alignments_t::VERTICAL_DISTRIBUTED`
   *
   * As in Excel, vertical and horizontal alignments can be combined:
   *
   * @code
   *  format = workbook.add_format();
   *
   *  format_set_align(format, xwpp::format_alignments_t::HORIZONTAL_CENTER);
   *  format_set_align(format, xwpp::format_alignments_t::VERTICAL_CENTER);
   *
   *  worksheet.set_row(0, 30);
   *  worksheet.write_string(0, 0, "Some Text", format);
   * @endcode
   *
   * @image html format_font_align.png
   *
   * Text can be aligned across two or more adjacent cells using the
   * `%center_across` property. However, for genuine merged cells it is better to
   * use the worksheet_t::merge_range() worksheet method.
   *
   * The vertical justify option can be used to provide automatic text wrapping
   * in a cell. The height of the cell will be adjusted to accommodate the
   * wrapped text. To specify where the text wraps use the
   * `set_text_wrap()` method.
   *
   * @todo Split in two functions or a two-paramters function.
   */
  void set_align(format_alignments_t alignment);

  /**
   * @brief Set the color of the font used in the cell.
   *
   * @param color The cell font color.
   *
   * Set the font color:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_font_color(xwpp::color_t::RED);
   *
   *  worksheet.write_string(0, 0, "Wheelbarrow", format);
   * @endcode
   *
   * @image html format_font_color.png
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   *
   * @note The `%set_font_color()` method is used to set the font color in a
   * cell. To set the color of a cell background use the `set_bg_color()`
   * and `set_pattern()` methods.
   */
  void set_font_color(color_t color);

  /**
   * @brief Set the pattern foreground color for a cell.
   *
   * @param color The cell pattern foreground  color.
   *
   * The `%set_fg_color()` method can be used to set the foreground color of
   * a pattern.
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_fg_color(color_t color);

  /**
   * @brief Set the pattern background color for a cell.
   *
   * @param color The cell pattern background color.
   *
   * The `%set_bg_color()` method can be used to set the background color of
   * a pattern. Patterns are defined via the `set_pattern()` method. If a
   * pattern hasn't been defined then a solid fill pattern is used as the
   * default.
   *
   * Here is an example of how to set up a solid fill in a cell:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_pattern(xwpp::format_patterns_t::SOLID);
   *  format->set_bg_color(xwpp::color_t::GREEN);
   *
   *  worksheet.write_string(0, 0, "Ray", format);
   * @endcode
   *
   * @image html formats_set_bg_color.png
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_bg_color(color_t color);

  /**
   * @brief Turn on underline for the format.
   *
   * @param style Underline style.
   *
   * Set the underline property of the format:
   *
   * @code
   *  format->set_underline(xwpp::format_underlines_t::SINGLE);
   * @endcode
   *
   * @image html format_font_underlined.png
   *
   * The available underline styles are:
   *
   * - `%format_underlines_t::SINGLE`
   * - `%format_underlines_t::DOUBLE`
   * - `%format_underlines_t::SINGLE_ACCOUNTING`
   * - `%format_underlines_t::DOUBLE_ACCOUNTING`
   *
   * @todo Check if overline is available on, if needed, add it.
   */
  void set_underline(format_underlines_t style);

  /**
   * @brief Wrap text in a cell.
   *
   * Turn text wrapping on for text in a cell.
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_text_wrap();
   *
   *  worksheet.write_string(0, 0, "Some long text to wrap in a cell", format);
   * @endcode
   *
   * If you wish to control where the text is wrapped you can add newline
   * characters to the string:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_text_wrap(format);
   *
   *  worksheet.write_string(0, 0, "It's\na bum\nwrap", format);
   * @endcode
   *
   * @image html format_font_text_wrap.png
   *
   * Excel will adjust the height of the row to accommodate the wrapped text. A
   * similar effect can be obtained without newlines using the
   * `set_align()` function with `%format_alignments_t::VERTICAL_JUSTIFY`.
   */
  void set_text_wrap();

  /**
   * @brief Set the cell border style.
   *
   * @param style Border style index.
   *
   * Set the cell border style:
   *
   * @code
   *  format->set_border(format_borders_t::THIN);
   * @endcode
   *
   * Individual border elements can be configured using the following functions
   * with the same parameters:
   *
   * - `set_bottom()`
   * - `set_top()`
   * - `set_left()`
   * - `set_right()`
   *
   * A cell border is comprised of a border on the bottom, top, left and right.
   * These can be set to the same value using `%set_border()` or
   * individually using the relevant method calls shown above.
   *
   * The following border styles are available:
   *
   * - `%format_borders_t::THIN`
   * - `%format_borders_t::MEDIUM`
   * - `%format_borders_t::DASHED`
   * - `%format_borders_t::DOTTED`
   * - `%format_borders_t::THICK`
   * - `%format_borders_t::DOUBLE`
   * - `%format_borders_t::HAIR`
   * - `%format_borders_t::MEDIUM_DASHED`
   * - `%format_borders_t::DASH_DOT`
   * - `%format_borders_t::MEDIUM_DASH_DOT`
   * - `%format_borders_t::DASH_DOT_DOT`
   * - `%format_borders_t::MEDIUM_DASH_DOT_DOT`
   * - `%format_borders_t::SLANT_DASH_DOT`
   *
   * The most commonly used style is the `thin` style.
   */
  void set_border(format_borders_t style);

  /**
   * @brief Set the cell bottom border style.
   *
   * @param style Border style index.
   *
   * Set the cell bottom border style.
   *
   * @see `set_border()` for details on the border styles.
   */
  void set_bottom(format_borders_t style);

  /**
   * @brief Set the cell top border style.
   *
   * @param style Border style index.
   *
   * Set the cell top border style.
   *
   * @see `set_border()` for details on the border styles.
   */
  void set_top(format_borders_t style);

  /**
   * @brief Set the cell left border style.
   *
   * @param style Border style index.
   *
   * Set the cell left border style.
   *
   * @see `set_border()` for details on the border styles.
   */
  void set_left(format_borders_t style);

  /**
   * @brief Set the cell right border style.
   *
   * @param style Border style index.
   *
   * Set the cell right border style.
   *
   * @see `set_border()` for details on the border styles.
   */
  void set_right(format_borders_t style);

  /**
   * @brief Set the background fill pattern for a cell.
   *
   * @param pattern Pattern index.
   *
   * Set the background pattern for a cell.
   *
   * The most common pattern is a solid fill of the background color:
   *
   * @code
   *  format = workbook.add_format();
   *
   *  format->set_pattern(xwpp::format_patterns_t::SOLID);
   *  format->set_bg_color(xwpp::color_t::YELLOW);
   * @endcode
   *
   * The available fill patterns are:
   *
   *    Fill Type                     | Define
   *    ----------------------------- | --------------------------------------
   *    Solid                         |  `%format_patterns_t::SOLID`
   *    Medium gray                   |  `%format_patterns_t::MEDIUM_GRAY`
   *    Dark gray                     |  `%format_patterns_t::DARK_GRAY`
   *    Light gray                    |  `%format_patterns_t::LIGHT_GRAY`
   *    Dark horizontal line          |  `%format_patterns_t::DARK_HORIZONTAL`
   *    Dark vertical line            |  `%format_patterns_t::DARK_VERTICAL`
   *    Dark diagonal stripe          |  `%format_patterns_t::DARK_DOWN`
   *    Reverse dark diagonal stripe  |  `%format_patterns_t::DARK_UP`
   *    Dark grid                     |  `%format_patterns_t::DARK_GRID`
   *    Dark trellis                  |  `%format_patterns_t::DARK_TRELLIS`
   *    Light horizontal line         |  `%format_patterns_t::LIGHT_HORIZONTAL`
   *    Light vertical line           |  `%format_patterns_t::LIGHT_VERTICAL`
   *    Light diagonal stripe         |  `%format_patterns_t::LIGHT_DOWN`
   *    Reverse light diagonal stripe |  `%format_patterns_t::LIGHT_UP`
   *    Light grid                    |  `%format_patterns_t::LIGHT_GRID`
   *    Light trellis                 |  `%format_patterns_t::LIGHT_TRELLIS`
   *    12.5% gray                    |  `%format_patterns_t::GRAY_125`
   *    6.25% gray                    |  `%format_patterns_t::GRAY_0625`
   */
  void set_pattern(format_patterns_t pattern);

  /**
   * @brief Set the color of the cell border.
   *
   * @param color The cell border color.
   *
   * Individual border elements can be configured using the following methods with
   * the same parameters:
   *
   * - `set_bottom_color()`
   * - `set_top_color()`
   * - `set_left_color()`
   * - `set_right_color()`
   *
   * Set the color of the cell borders. A cell border is comprised of a border
   * on the bottom, top, left and right. These can be set to the same color
   * using `set_border_color()` or individually using the relevant method
   * calls shown above.
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_border_color(color_t color);

  /**
   * @brief Set the color of the bottom cell border.
   *
   * @param color The cell border color.
   *
   * @see `set_border_color()` for details on the border colors.
   */
  void set_bottom_color(color_t color);

  /**
   * @brief Set the color of the top cell border.
   *
   * @param color The cell border color.
   *
   * @see `set_border_color()` for details on the border colors.
   */
  void set_top_color(color_t color);

  /**
   * @brief Set the color of the left cell border.
   *
   * @param color The cell border color.
   *
   * @see `set_border_color()` for details on the border colors.
   */
  void set_left_color(color_t color);

  /**
   * @brief Set the color of the right cell border.
   *
   * @param color The cell border color.
   *
   * @see `set_border_color()` for details on the border colors.
   */
  void set_right_color(color_t color);

  /**
   * @brief Set the diagonal cell border type.
   *
   * @param type The `format_diagonal_types` diagonal border type.
   *
   * Set the diagonal cell border type:
   *
   * @code
   *  xwpp::format_t* format1 = workbook.add_format();
   *  format1->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP);
   *
   *  xwpp::format_t* format2 = workbook.add_format();
   *  format2->set_diag_type(xwpp::format_diagonal_types_t::BORDER_DOWN);
   *
   *  xwpp::format_t* format3 = workbook.add_format();
   *  format3->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN);
   *
   *  xwpp::format_t* format4 = workbook.add_format();
   *  format4->set_diag_type(xwpp::format_diagonal_types_t::BORDER_UP_DOWN);
   *  format4->set_diag_border(xwpp::format_borders_t::HAIR);
   *  format4->set_diag_color(xwpp::color_t::RED);
   *
   *  worksheet.write_string(CELL("B3"), "Text", format1);
   *  worksheet.write_string(CELL("B6"), "Text", format2);
   *  worksheet.write_string(CELL("B9"), "Text", format3);
   *  worksheet.write_string(CELL("B12"), "Text", format4);
   * @endcode
   *
   * @image html diagonal_border.png
   *
   * The allowable border types are defined in `format_diagonal_types_t`:
   *
   * - `%format_diagonal_types_t::BORDER_UP`: Cell diagonal border from bottom left to top
   *   right.
   *
   * - `%format_diagonal_types_t::BORDER_DOWN`: Cell diagonal border from top left to bottom
   *   right.
   *
   * - `%format_diagonal_types_t::BORDER_UP_DOWN`: Cell diagonal border from top left to
   *   bottom right. A combination of the 2 previous types.
   *
   * If the border style isn't specified with `set_diag_border()` then it
   * will default to `format_borders_t::THIN`.
   */
  void set_diag_type(format_diagonal_types_t type);

  /**
   * @brief Set the diagonal cell border style.
   *
   * @param style The `format_borders_t` style.
   *
   * Set the diagonal border style. This should be a `format_borders_t` value.
   */
  void set_diag_border(format_borders_t style);

  /**
   * @brief Set the diagonal cell border color.
   *
   * @param color The cell diagonal border color.
   *
   * Set the diagonal border color. The color should be an RGB integer value,
   *
   * @see @ref working_with_colors and the above example.
   */
  void set_diag_color(color_t color);

  /**
   * @brief Set the Excel built-in number format for a cell.
   *
   * @param index The built-in number format index for the cell.
   *
   * This function is similar to `%set_num_format()` except that it takes an
   * index to a limited number of Excel's built-in number formats instead of a
   * user defined format string:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_num_format_index(0x0F); // d-mmm-yy
   * @endcode
   *
   * @note Unless you need to specifically access one of Excel's built-in number
   * formats the `set_num_format()` function above is a better solution.
   * The `%set_num_format_index()` function is mainly included for
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
   *  - Numeric formats 23 to 36 are not documented by Microsoft and may differ
   *    in international versions. The listed date and currency formats may also
   *    vary depending on system settings.
   *  - The dollar sign in the above format appears as the defined local currency
   *    symbol.
   *  - These formats can also be set via `set_num_format()`.
   *
   * @see also @ref ww_formats_categories.
   *
   * @todo Use dedicated enum.
   */
  void set_num_format_index(uint8_t index);

  /**
   * @brief Set the font used in the cell.
   *
   * @param font_name Cell font name.
   *
   * Specify the font used used in the cell format:
   *
   * @code
   *  format->set_font_name("Avenir Black Oblique");
   * @endcode
   *
   * @image html format_set_font_name.png
   *
   * Excel can only display fonts that are installed on the system that it is
   * running on. Therefore it is generally best to use the fonts that come as
   * standard with Excel such as Calibri, Times New Roman and Courier New.
   *
   * The default font in Excel 2007, and later, is `Calibri`.
   */
  void set_font_name(const std::string& font_name);

  /**
   * @brief Set the size of the font used in the cell.
   *
   * @param size The cell font size.
   *
   * Set the font size of the cell format:
   *
   * @code
   *  format->set_font_size(30);
   * @endcode
   *
   * @image html format_font_size.png
   *
   * Excel adjusts the height of a row to accommodate the largest font
   * size in the row. You can also explicitly specify the height of a
   * row using the `worksheet_t::set_row()` function.
   */
  void set_font_size(double size);

  /**
   * @brief Set the strikeout property of the font.
   *
   * @image html format_font_strikeout.png
   *
   * @todo Check if there is strike type in Excel. And
   * add parameter if available.
   */
  void set_font_strikeout();

  /**
   * @brief Set the superscript/subscript property of the font.
   *
   * @param style Superscript or subscript style.
   *
   * Set the superscript or subscript property of the font.
   *
   * @image html format_font_script.png
   *
   * The available script styles are:
   *
   * - `format_scripts_t::SUPERSCRIPT`
   * - `format_scripts_t::SUBSCRIPT`
   */
  void set_font_script(format_scripts_t style);

  /**
   * @brief Set the Format font family property.
   *
   * @param value The font family index.
   *
   * Set the font family. This is usually an integer in the range 1-4.
   *
   * @code
   *  format->set_font_family(178);
   * @endcode
   */
  void set_font_family(uint8_t value);

  /**
   * @brief Set the Format font character set property.
   *
   * @param value The font character set.
   *
   * Set the font character set property.
   *
   * @code
   *  format->set_font_charset(178);
   * @endcode
   */
  void set_font_charset(uint8_t value);

  /**
   * @brief Set the Format font outline property.
   *
   * Set the font outline property.
   *
   * @code
   *  format->set_font_outline();
   * @endcode
   *
   * @todo Add image in description.
   */
  void set_font_outline();

  /**
   * @brief Set the Format font shadow property.
   *
   * Set the font shadow property.
   *
   * @code
   *  format->set_font_shadow();
   * @endcode
   *
   * @todo Add image in description.
   */
  void set_font_shadow();

  /**
   * @brief Set the Format font scheme property.
   *
   * @param font_scheme The font scheme.
   *
   * Set the font scheme property.
   *
   * @code
   *  format->set_font_scheme("fs");
   * @endcode
   *
   * @todo Add image in description.
   * @todo Add an example (and fix example here).
   */
  void set_font_scheme(const std::string& font_scheme);

  /**
   * @brief Set the Format font condense property.
   *
   * Set the font condense property.
   *
   * @code
   *  format->set_font_condense();
   * @endcode
   *
   * @todo Add image in description.
   */
  void set_font_condense();

  /**
   * @brief Set the Format font extend property.
   *
   * Set the font extend property.
   *
   * @code
   *  format->set_font_extend();
   * @endcode
   *
   * @todo Add image in description.
   */
  void set_font_extend();

  /**
   * @brief Set the Format font only property.
   *
   * Set the font only property.
   *
   * @code
   *  format->set_font_only();
   * @endcode
   *
   * @todo Add image in description.
   * @todo Add an example (and fix example here).
   */
  void set_font_only();

  /**
   * @brief Set the cell unlocked state.
   *
   * This property can be used to allow modification of a cell in a protected
   * worksheet. In Excel, cell locking is turned on by default for all
   * cells. However, it only has an effect if the worksheet has been protected
   * using the worksheet `worksheet_t::protect()` function:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_unlocked();
   *
   *  // Enable worksheet protection, without password or options.
   *  worksheet.protect();
   *
   *  // This cell cannot be edited.
   *  worksheet.write_formula(0, 0, "=1+2");
   *
   *  // This cell can be edited.
   *  worksheet.write_formula(1, 0, "=1+2", format);
   * @endcode
   */
  void set_unlocked();

  /**
   * @brief Hide formulas in a cell.
   *
   * This property is used to hide a formula while still displaying its
   * result. This is generally used to hide complex calculations from end users
   * who are only interested in the result. It only has an effect if the
   * worksheet has been protected using the worksheet `worksheet_t::protect()`
   * function:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_hidden();
   *
   *  // Enable worksheet protection, without password or options.
   *  worksheet.protect();
   *
   *  // The formula in this cell isn't visible.
   *  worksheet.write_formula(0, 0, "=1+2", format);
   * @endcode
   */
  void set_hidden();

  /**
   * @brief Set the rotation of the text in a cell.
   *
   * @param angle Rotation angle in the range -90 to 90 and 270.
   *
   * Set the rotation of the text in a cell. The rotation can be any angle in the
   * range -90 to 90 degrees:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_rotation(30);
   *
   *  worksheet.write_string(0, 0, "This text is rotated", format);
   * @endcode
   *
   * @image html format_font_text_rotated.png
   *
   * The angle 270 is also supported. This indicates text where the letters run
   * from top to bottom.
   *
   * @todo Add specific API for 270.
   */
  void set_rotation(int16_t angle);

  /**
   * @brief Set the cell text indentation level.
   *
   * @param level Indentation level.
   *
   * This method can be used to indent text in a cell. The argument, which should
   * be an integer, is taken as the level of indentation:
   *
   * @code
   *  format1 = workbook.add_format();
   *  format2 = workbook.add_format();
   *
   *  format1->set_indent(1);
   *  format2->set_indent(2);
   *
   *  worksheet.write_string(0, 0, "This text is indented 1 level", format1);
   *  worksheet_write_string(1, 0, "This text is indented 2 levels", format2);
   * @endcode
   *
   * @image html text_indent.png
   *
   * @note Indentation is a horizontal alignment property. It will override any other
   * horizontal properties but it can be used in conjunction with vertical
   * properties.
   */
  void set_indent(uint8_t level);

  /**
   * @brief Turn on the text "shrink to fit" for a cell.
   *
   * This method can be used to shrink text so that it fits in a cell:
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_shrink();
   *
   *  worksheet.write_string(0, 0, "Honey, I shrunk the text!", format);
   * @endcode
   */
  void set_shrink();

  // TODO Useless, to check (missing parameter? Remove?)
  void set_text_justlast();

  /**
   * @brief Set the vertical alignment.
   *
   * @param alignment A `format_alignments_t` indicating the vertical alignment.
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_valign(xwpp::format_alignments_t::VERTICAL_BOTTOM);
   *
   *  worksheet.write_string(0, 0, "Bottom alignment", format);
   * @endcode
   *
   * @todo Add image in documentation.
   * @todo Add example.
   */
  void set_valign(format_alignments_t alignment);

  /**
   * @brief Set the horizontal alignment.
   *
   * @param alignment A `format_alignments_t` indicating the horizontal alignment.
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_halign(xwpp::format_alignments_t::HORIZONTAL_RIGHT);
   *
   *  worksheet.write_string(0, 0, "Right alignment", format);
   * @endcode
   *
   * @todo Add image in documentation.
   * @todo Add example.
   */
  void set_halign(format_alignments_t alignment);

  /**
   * @brief Set reading order.
   *
   * @param value The reading order.
   *
   * Set the reading order.
   *
   * @todo Add image in documentation.
   * @todo Add enum for reading order.
   */
  void set_reading_order(uint8_t value);

  /**
   * @brief Set the theme.
   *
   * @param value The theme.
   *
   * Set the theme.
   *
   * @todo Add image in documentation.
   * @todo Add enum for theme.
   * @todo Add example.
   */
  void set_theme(uint8_t value);

  /**
   * @brief Set the indexed color.
   *
   * @param value The indexed color.
   *
   * Set the indexed color.
   *
   * @todo Add image in documentation.
   * @todo Add enum for theme.
   * @todo Add example.
   */
  void set_color_indexed(uint8_t value);

  /**
   * @brief Turn on quote prefix for the format.
   *
   * Set the quote prefix property of a format to ensure a string is treated
   * as a string after editing. This is the same as prefixing the string with
   * a single quote in Excel. You don't need to add the quote to the
   * string but you do need to add the format.
   *
   * @code
   *  format = workbook.add_format();
   *  format->set_quote_prefix();
   *
   *  worksheet.write_string(0, 0, "=Foo", format);
   * @endcode
   */
  void set_quote_prefix();

  /**
   * @brief Set the format as hyperlink format.
   *
   * Set the format as hyperlink format.
   *
   * @todo Add image.
   */
  void set_hyperlink();

  static const int32_t PROPERTY_UNSET = -1;
  static const std::string DEFAULT_FONT_NAME;

  // TODO Refactor this point to not have it in several places
  // NOLINTNEXTLINE(cppcoreguidelines-non-private-member-variables-in-classes,misc-non-private-member-variables-in-classes)
  std::function<int32_t(format_t*)> get_dxf_index_;

private:
  // TODO friend up to refactoring with struct / class
  friend class workbook_t;
  friend class worksheet_t; // TODO for debug
  friend class style_t;

  static const uint8_t DEFAULT_FONT_FAMILY = 2;

  int32_t xf_index_  = PROPERTY_UNSET;
  int32_t dxf_index_ = PROPERTY_UNSET;
  int32_t xf_id_     = 0;
  std::string num_format_;
  std::string font_name_;
  std::string font_scheme_;
  uint16_t num_format_index_         = 0;
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
  uint8_t font_charset_              = 0;
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

/// @cond
}
/// @endcond

#endif
