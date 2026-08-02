/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page worksheet_page The Worksheet object
 *
 * @see @ref worksheet.h for full details of the functionality.
 *
 * @file
 *
 * @brief Management of an Excel worksheet.
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * @code
 *  #include "xlsxwriterpp.h"
 *
 *  int main()
 *  {
 *    xwpp::workbook_t workbook;
 *    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
 *
 *    worksheet.write_string(0, 0, "Hello Excel");
 *
 *    workbook.save("filename.xlsx");
 *  }
 * @endcode
 */

#ifndef XWPP_WORKSHEET_H
#define XWPP_WORKSHEET_H

#include "xwpp/chart.h"
#include "xwpp/common.h"
#include "xwpp/drawing.h"
#include "xwpp/format.h"
#include "xwpp/shared_strings.h"

#include <chrono>
#include <cstdint>
#include <functional>
#include <map>
#include <optional>
#include <string>
#include <variant>

/// @cond
namespace xwpp
{
/// @endcond

const uint32_t HEADER_FOOTER_MAX = 255;

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
const size_t BREAKS_MAX = 1023;

/** Default Excel column width in character units. */
const double DEF_COL_WIDTH = 8.43;

/** Default Excel row height in character units. */
const double DEF_ROW_HEIGHT = 15.0;

/** Default Excel column width in pixels. */
const uint32_t DEF_COL_WIDTH_PIXELS = 64;

/** Default Excel column height in pixels. */
const uint32_t DEF_ROW_HEIGHT_PIXELS = 20;

// Conversion functions.
double pixels_to_height(double pixels);
double pixels_to_width(double pixels);

/**
 * @brief Gridline options.
 *
 * Gridline options using in `worksheet_t::gridlines()`.
 *
 * @todo Use `enum class`.
 */
enum gridlines_t
{
  /** Hide screen and print gridlines. */
  HIDE_ALL_GRIDLINES = 0,

  /** Show screen gridlines. */
  SHOW_SCREEN_GRIDLINES = 1,

  /** Show print gridlines. */
  SHOW_PRINT_GRIDLINES = 2,

  /** Show screen and print gridlines. */
  SHOW_ALL_GRIDLINES = 3,
};

/**
 * @brief Data validation property values.
 *
 * Data validation property values.
 */
enum class validation_boolean_t
{
  DEFAULT,

  /** Turn a data validation property off. */
  OFF,

  /**
   * Turn a data validation property on.
   *
   * Data validation properties are generally on by default.
   */
  ON
};

/**
 * @brief Data validation types.
 *
 * Data validation types.
 */
enum class validation_types_t
{
  NONE,

  /** Restrict cell input to whole/integer numbers only. */
  INTEGER,

  /** Restrict cell input to whole/integer numbers only, using a cell
   *  reference. */
  INTEGER_FORMULA,

  /** Restrict cell input to decimal numbers only. */
  DECIMAL,

  /** Restrict cell input to decimal numbers only, using a cell
   * reference. */
  DECIMAL_FORMULA,

  /** Restrict cell input to a list of strings in a dropdown. */
  LIST,

  /** Restrict cell input to a list of strings in a dropdown, using a
   * cell range. */
  LIST_FORMULA,

  /** Restrict cell input to date values only, using a lxw_datetime type. */
  DATE,

  /** Restrict cell input to date values only, using a cell reference. */
  DATE_FORMULA,

  /* Restrict cell input to date values only, as a serial number.
   * Undocumented. */
  DATE_NUMBER,

  /** Restrict cell input to time values only, using a lxw_datetime type. */
  TIME,

  /** Restrict cell input to time values only, using a cell reference. */
  TIME_FORMULA,

  /* Restrict cell input to time values only, as a serial number.
   * Undocumented. */
  TIME_NUMBER,

  /** Restrict cell input to strings of defined length, using a cell
   * reference. */
  LENGTH,

  /** Restrict cell input to strings of defined length, using a cell
   * reference. */
  LENGTH_FORMULA,

  /** Restrict cell to input controlled by a custom formula that returns
   * `TRUE/FALSE`. */
  CUSTOM_FORMULA,

  /** Allow any type of input. Mainly only useful for pop-up messages. */
  ANY
};

/**
 * @brief Data validation criteria.
 *
 * Data validation criteria uses to control the selection of data.
 */
enum class validation_criteria_t
{
  NONE,

  /** Select data between two values. */
  BETWEEN,

  /** Select data that is not between two values. */
  NOT_BETWEEN,

  /** Select data equal to a value. */
  EQUAL_TO,

  /** Select data not equal to a value. */
  NOT_EQUAL_TO,

  /** Select data greater than a value. */
  GREATER_THAN,

  /** Select data less than a value. */
  LESS_THAN,

  /** Select data greater than or equal to a value. */
  GREATER_THAN_OR_EQUAL_TO,

  /** Select data less than or equal to a value. */
  LESS_THAN_OR_EQUAL_TO
};

/**
 * @brief Data validation error types.
 *
 * Data validation error types for pop-up messages.
 */
enum class validation_error_types_t
{
  /**
   * Show a "Stop" data validation pop-up message.
   *
   * This is the default.
   */
  STOP,

  /** Show an "Error" data validation pop-up message. */
  WARNING,

  /** Show an "Information" data validation pop-up message. */
  INFORMATION
};

/**
 * @brief Display type for a cell comment.
 *
 * Set the display type for a cell comment. This is hidden by default but
 * can be set to visible with the `worksheet_t::show_comments()` function.
 */
enum class comment_display_t
{
  /** Default to the worksheet default which can be hidden or visible. */
  DEFAULT,

  /**
   * Hide the cell comment.
   *
   * Usually the default.
   */
  HIDDEN,

  /**
   * Show the cell comment.
   *
   * Can also be set for the worksheet with the `worksheet_t::show_comments()` function.
   */
  VISIBLE
};

/**
 * @brief Type definitions for conditional formats.
 *
 * Values used to set the "type" field of conditional format.
 */
enum class conditional_format_types_t
{
  NONE,

  /** The Cell type is the most common conditional formatting type. It is
   *  used when a format is applied to a cell based on a simple
   *  criterion.  */
  CELL,

  /** The Text type is used to specify Excel's "Specific Text" style
   *  conditional format. */
  TEXT,

  /** The Time Period type is used to specify Excel's "Dates Occurring"
   *  style conditional format. */
  TIME_PERIOD,

  /** The Average type is used to specify Excel's "Average" style
   *  conditional format. */
  AVERAGE,

  /** The Duplicate type is used to highlight duplicate cells in a range. */
  DUPLICATE,

  /** The Unique type is used to highlight unique cells in a range. */
  UNIQUE,

  /** The Top type is used to specify the top n values by number or
   *  percentage in a range. */
  TOP,

  /** The Bottom type is used to specify the bottom n values by number or
   *  percentage in a range. */
  BOTTOM,

  /** The Blanks type is used to highlight blank cells in a range. */
  BLANKS,

  /** The No Blanks type is used to highlight non blank cells in a range. */
  NO_BLANKS,

  /** The Errors type is used to highlight error cells in a range. */
  ERRORS,

  /** The No Errors type is used to highlight non error cells in a range. */
  NO_ERRORS,

  /** The Formula type is used to specify a conditional format based on a
   *  user defined formula. */
  FORMULA,

  /** The 2 Color Scale type is used to specify Excel's "2 Color Scale"
   *  style conditional format. */
  TWO_COLOR_SCALE,

  /** The 3 Color Scale type is used to specify Excel's "3 Color Scale"
   *  style conditional format. */
  THREE_COLOR_SCALE,

  /** The Data Bar type is used to specify Excel's "Data Bar" style
   *  conditional format. */
  DATA_BAR,

  /** The Icon Set type is used to specify a conditional format with a set
   *  of icons such as traffic lights or arrows. */
  ICON_SETS,
};

/**
 * @brief The criteria used in a conditional format.
 *
 * Criteria used to define how a conditional format works.
 */
enum class conditional_criteria_t
{
  NONE,

  /** Format cells equal to a value. */
  EQUAL_TO,

  /** Format cells not equal to a value. */
  NOT_EQUAL_TO,

  /** Format cells greater than a value. */
  GREATER_THAN,

  /** Format cells less than a value. */
  LESS_THAN,

  /** Format cells greater than or equal to a value. */
  GREATER_THAN_OR_EQUAL_TO,

  /** Format cells less than or equal to a value. */
  LESS_THAN_OR_EQUAL_TO,

  /** Format cells between two values. */
  BETWEEN,

  /** Format cells that is not between two values. */
  NOT_BETWEEN,

  /** Format cells that contain the specified text. */
  TEXT_CONTAINING,

  /** Format cells that don't contain the specified text. */
  TEXT_NOT_CONTAINING,

  /** Format cells that begin with the specified text. */
  TEXT_BEGINS_WITH,

  /** Format cells that end with the specified text. */
  TEXT_ENDS_WITH,

  /** Format cells with a date of yesterday. */
  TIME_PERIOD_YESTERDAY,

  /** Format cells with a date of today. */
  TIME_PERIOD_TODAY,

  /** Format cells with a date of tomorrow. */
  TIME_PERIOD_TOMORROW,

  /** Format cells with a date in the last 7 days. */
  TIME_PERIOD_LAST_7_DAYS,

  /** Format cells with a date in the last week. */
  TIME_PERIOD_LAST_WEEK,

  /** Format cells with a date in the current week. */
  TIME_PERIOD_THIS_WEEK,

  /** Format cells with a date in the next week. */
  TIME_PERIOD_NEXT_WEEK,

  /** Format cells with a date in the last month. */
  TIME_PERIOD_LAST_MONTH,

  /** Format cells with a date in the current month. */
  TIME_PERIOD_THIS_MONTH,

  /** Format cells with a date in the next month. */
  TIME_PERIOD_NEXT_MONTH,

  /** Format cells above the average for the range. */
  AVERAGE_ABOVE,

  /** Format cells below the average for the range. */
  AVERAGE_BELOW,

  /** Format cells above or equal to the average for the range. */
  AVERAGE_ABOVE_OR_EQUAL,

  /** Format cells below or equal to the average for the range. */
  AVERAGE_BELOW_OR_EQUAL,

  /** Format cells 1 standard deviation above the average for the range. */
  AVERAGE_1_STD_DEV_ABOVE,

  /** Format cells 1 standard deviation below the average for the range. */
  AVERAGE_1_STD_DEV_BELOW,

  /** Format cells 2 standard deviation above the average for the range. */
  AVERAGE_2_STD_DEV_ABOVE,

  /** Format cells 2 standard deviation below the average for the range. */
  AVERAGE_2_STD_DEV_BELOW,

  /** Format cells 3 standard deviation above the average for the range. */
  AVERAGE_3_STD_DEV_ABOVE,

  /** Format cells 3 standard deviation below the average for the range. */
  AVERAGE_3_STD_DEV_BELOW,

  /** Format cells in the top of bottom percentage. */
  TOP_OR_BOTTOM_PERCENT
};

/**
 * @brief Conditional format rule types.
 *
 * Conditional format rule types that apply to Color Scale and Data Bars.
 */
enum class conditional_format_rule_types_t
{
  NONE,

  /** Conditional format rule type: matches the minimum values in the
   *  range.
   *
   * Can only be applied to `min_rule_type_`.*/
  MINIMUM,

  /** Conditional format rule type: use a number to set the bound.*/
  NUMBER,

  /** Conditional format rule type: use a percentage to set the bound.*/
  PERCENT,

  /** Conditional format rule type: use a percentile to set the bound.*/
  PERCENTILE,

  /** Conditional format rule type: use a formula to set the bound.*/
  FORMULA,

  /** Conditional format rule type: matches the maximum values in the
   *  range.
   *
   * Can only be applied to max_rule_type.*/
  MAXIMUM,

  /* Used internally for Excel2010 bars. Not documented. */
  AUTO_MIN,

  /* Used internally for Excel2010 bars. Not documented. */
  AUTO_MAX
};

/**
 * @brief Conditional format data bar directions.
 *
 * Values used to set the bar direction of a conditional format data bar.
 */
enum class conditional_format_bar_direction_t
{

  /** Data bar direction is set by Excel based on the context of the data
   *  displayed. */
  CONTEXT,

  /** Data bar direction is from right to left. */
  RIGHT_TO_LEFT,

  /** Data bar direction is from left to right. */
  LEFT_TO_RIGHT
};

/**
 * @brief Conditional format data bar axis options.
 *
 * Values used to set the position of the axis in a conditional format data
 * bar.
 */
enum class conditional_bar_axis_position_t
{

  /** Data bar axis position is set by Excel based on the context of the
   *  data displayed. */
  AUTOMATIC,

  /** Data bar axis position is set at the midpoint. */
  MIDPOINT,

  /** Data bar axis is turned off. */
  NONE
};

/**
 * @brief Icon types used in the `conditional_format_t` `icon_style_` field.
 *
 * Definitions of icon styles used with Icon Set conditional formats.
 */
enum class conditional_icon_types_t
{
  /** 3 colored arrows showing up, sideways and down. */
  THREE_ARROWS_COLORED,

  /** 3 gray arrows showing up, sideways and down. */
  THREE_ARROWS_GRAY,

  /** 3 colored flags in red, yellow and green. */
  THREE_FLAGS,

  /** 3 traffic lights - rounded. */
  THREE_TRAFFIC_LIGHTS_UNRIMMED,

  /** 3 traffic lights with a rim - squarish. */
  THREE_TRAFFIC_LIGHTS_RIMMED,

  /** 3 colored shapes - a circle, triangle and diamond. */
  THREE_SIGNS,

  /** 3 circled symbols with tick mark, exclamation and cross. */
  THREE_SYMBOLS_CIRCLED,

  /** 3 symbols with tick mark, exclamation and cross. */
  THREE_SYMBOLS_UNCIRCLED,

  /** 4 colored arrows showing up, diagonal up, diagonal down and down. */
  FOUR_ARROWS_COLORED,

  /** 4 gray arrows showing up, diagonal up, diagonal down and down. */
  FOUR_ARROWS_GRAY,

  /** 4 circles in 4 colors going from red to black. */
  FOUR_RED_TO_BLACK,

  /** 4 histogram ratings. */
  FOUR_RATINGS,

  /** 4 traffic lights. */
  FOUR_TRAFFIC_LIGHTS,

  /** 5 colored arrows showing up, diagonal up, sideways, diagonal down and down. */
  FIVE_ARROWS_COLORED,

  /** 5 gray arrows showing up, diagonal up, sideways, diagonal down and down. */
  FIVE_ARROWS_GRAY,

  /** 5 histogram ratings. */
  FIVE_RATINGS,

  /** 5 quarters, from 0 to 4 quadrants filled. */
  FIVE_QUARTERS
};

/**
 * @brief The type of table style.
 *
 * The type of table style.
 */
enum class table_style_type_t
{
  DEFAULT,

  /** Light table style. */
  LIGHT,

  /** Medium table style. */
  MEDIUM,

  /** Dark table style. */
  DARK
};

/**
 * @brief Standard Excel functions for totals in tables.
 *
 * Definitions for the standard Excel functions that are available via the
 * dropdown in the total row of an Excel table.
 */
enum class table_total_functions_t
{
  NONE = 0,

  /** Use the average function as the table total. */
  AVERAGE = 101,

  /** Use the count numbers function as the table total. */
  COUNT_NUMS = 102,

  /** Use the count function as the table total. */
  COUNT = 103,

  /** Use the max function as the table total. */
  MAX = 104,

  /** Use the min function as the table total. */
  MIN = 105,

  /** Use the standard deviation function as the table total. */
  STD_DEV = 107,

  /** Use the sum function as the table total. */
  SUM = 109,

  /** Use the var function as the table total. */
  VAR = 110
};

/**
 * @brief The criteria used in autofilter rules.
 *
 * Criteria used to define an autofilter rule condition.
 */
enum class filter_criteria_t
{
  NONE,

  /** Filter cells equal to a value. */
  EQUAL_TO,

  /** Filter cells not equal to a value. */
  NOT_EQUAL_TO,

  /** Filter cells greater than a value. */
  GREATER_THAN,

  /** Filter cells less than a value. */
  LESS_THAN,

  /** Filter cells greater than or equal to a value. */
  GREATER_THAN_OR_EQUAL_TO,

  /** Filter cells less than or equal to a value. */
  LESS_THAN_OR_EQUAL_TO,

  /** Filter cells that are blank. */
  BLANKS,

  /** Filter cells that are not blank. */
  NON_BLANKS
};

/**
 * @brief And/or operator when using 2 filter rules.
 *
 * And/or operator conditions when using 2 filter rules with
 * `worksheet_t::filter_column2()`.
 *
 * In general `filter_operator_t::OR` is used with `filter_criteria_t::EQUAL_TO` and
 * `filter_operator_t::AND` is used with the other filter criteria.
 */
enum class filter_operator_t
{
  /** Logical "and" of 2 filter rules. */
  AND,

  /** Logical "or" of 2 filter rules. */
  OR
};

enum class filter_type_t
{
  NONE,
  SINGLE,
  AND,
  OR,
  STRING_LIST
};

/**
 * @brief Options to control the positioning of objects.
 *
 * Options to control the positioning of worksheet objects such as images
 * or charts. See @ref working_with_object_positioning.
 */
enum class object_position_t
{
  /** Default positioning for the object. */
  DEFAULT,

  /** Move and size the worksheet object with the cells. */
  MOVE_AND_SIZE,

  /** Move but don't size the worksheet object with the cells. */
  MOVE_DONT_SIZE,

  /** Don't move or size the worksheet object with the cells. */
  DONT_MOVE_DONT_SIZE,

  /** Same as `MOVE_AND_SIZE` except `Xlsxwriter++` applies hidden
   *  cells after the object is inserted. */
  MOVE_AND_SIZE_AFTER
};

/**
 * @brief Options for ignoring errors/warnings.
 *
 * Options for ignoring worksheet errors/warnings. See
 * `worksheet_t::ignore_errors()`.
 */
enum class ignore_errors_t
{
  /** Turn off errors/warnings for numbers stores as text. */
  NUMBER_STORED_AS_TEXT = 1,

  /** Turn off errors/warnings for formula errors (such as divide by
   *  zero). */
  EVAL_ERROR,

  /** Turn off errors/warnings for formulas that differ from surrounding
   *  formulas. */
  FORMULA_DIFFERS,

  /** Turn off errors/warnings for formulas that omit cells in a range. */
  FORMULA_RANGE,

  /** Turn off errors/warnings for unlocked cells that contain formulas. */
  FORMULA_UNLOCKED,

  /** Turn off errors/warnings for formulas that refer to empty cells. */
  EMPTY_CELL_REFERENCE,

  /** Turn off errors/warnings for cells in a table that do not comply with
   *  applicable data validation rules. */
  LIST_DATA_VALIDATION,

  /** Turn off errors/warnings for cell formulas that differ from the column
   *  formula. */
  CALCULATED_COLUMN,

  /** Turn off errors/warnings for formulas that contain a two digit text
   *  representation of a year. */
  TWO_DIGIT_TEXT_YEAR,

  LAST_OPTION
};

enum class cell_types_t
{
  NUMBER_CELL = 1,
  STRING_CELL,
  FORMULA_CELL,
  ARRAY_FORMULA_CELL,
  DYNAMIC_ARRAY_FORMULA_CELL,
  BLANK_CELL,
  BOOLEAN_CELL,
  ERROR_CELL,
  COMMENT,
  HYPERLINK_URL,
  HYPERLINK_INTERNAL,
  HYPERLINK_EXTERNAL
};

enum class pane_types_t
{
  NO_PANES = 0,
  FREEZE_PANES,
  SPLIT_PANES,
  FREEZE_SPLIT_PANES
};

enum class image_position_t
{
  HEADER_LEFT = 0,
  HEADER_CENTER,
  HEADER_RIGHT,
  FOOTER_LEFT,
  FOOTER_CENTER,
  FOOTER_RIGHT
};

// Internal structure for VML object options.
struct vml_obj_t
{
  row_num_t row_;
  col_num_t col_;
  row_num_t start_row_;
  col_num_t start_col_;
  int32_t x_offset_;
  int32_t y_offset_;
  uint64_t col_absolute_;
  uint64_t row_absolute_;
  uint32_t width_;
  uint32_t height_;
  double x_dpi_;
  double y_dpi_;
  color_t color_ = color_t::UNSET;
  uint8_t font_family_;
  comment_display_t visible_ = comment_display_t::DEFAULT;
  uint32_t author_id_        = 0;
  uint32_t rel_index_;
  double font_size_;
  drawing_coords_t from_;
  drawing_coords_t to_;
  std::string author_;
  std::string font_name_;
  std::string text_;
  std::string image_position_;
  std::string name_;
  std::string macro_;
};

struct cell_t
{
  row_num_t row_num_ = 0;
  col_num_t col_num_ = 0;
  cell_types_t type_ = cell_types_t::BLANK_CELL;
  format_t* format_  = nullptr;
  std::optional<vml_obj_t> comment_;
  std::variant<uint32_t, double, std::string> data_;
  double formula_result_;
  std::string user_data1_;
  std::string user_data2_;
  std::string sst_string_;
};

struct row_t
{
  row_num_t row_num_   = 0;
  double height_       = DEF_ROW_HEIGHT;
  format_t* format_    = nullptr;
  bool hidden_         = false;
  uint8_t level_       = 0;
  bool collapsed_      = false;
  bool row_changed_    = false;
  bool data_changed_   = false;
  bool height_changed_ = false;

  std::map<col_num_t, cell_t> cells_;
};

struct table_rows_t
{
  row_t& get_row_list(row_num_t row_num);

  std::map<row_num_t, row_t> rbh_root_; // NOLINT(misc-non-private-member-variables-in-classes)
};

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the `worksheet_t::set_column()` and
 * `worksheet_t::set_row()` functions.
 *
 * The members of this struct are explained in @ref ww_outlines_grouping.
 */
struct row_col_options_t
{
  /**
   * @brief Hide the row/column.
   *
   * Hide the row/column.
   */
  bool hidden_ = false;

  /**
   * @brief Outline level.
   *
   * Outline level.
   */
  uint8_t level_ = 0;

  /**
   * @brief Set the outline row as collapsed.
   *
   * Set the outline row as collapsed.
   */
  bool collapsed_ = false;
};

struct col_options_t
{
  col_num_t firstcol_ = std::numeric_limits<col_num_t>::max();
  col_num_t lastcol_  = std::numeric_limits<col_num_t>::max();
  double width_       = DEF_COL_WIDTH;
  format_t* format_   = nullptr;
  bool hidden_        = false;
  uint8_t level_      = 0;
  bool collapsed_     = false;
};

struct merged_range_t
{
  row_num_t first_row_;
  row_num_t last_row_;
  col_num_t first_col_;
  col_num_t last_col_;
};

struct repeat_rows_t
{
  bool in_use_         = false;
  row_num_t first_row_ = 0;
  row_num_t last_row_  = 0;
};

struct repeat_cols_t
{
  bool in_use_         = false;
  col_num_t first_col_ = 0;
  col_num_t last_col_  = 0;
};

struct print_area_t
{
  bool in_use_         = false;
  row_num_t first_row_ = 0;
  row_num_t last_row_  = 0;
  col_num_t first_col_ = 0;
  col_num_t last_col_  = 0;
};

struct autofilter_t
{
  bool in_use_         = false;
  bool has_rules_      = false;
  row_num_t first_row_ = 0;
  row_num_t last_row_  = 0;
  col_num_t first_col_ = 0;
  col_num_t last_col_  = 0;
};

struct panes_t
{
  pane_types_t type_   = pane_types_t::NO_PANES;
  row_num_t first_row_ = 0;
  col_num_t first_col_ = 0;
  row_num_t top_row_   = 0;
  col_num_t left_col_  = 0;
  double x_split_      = 0.;
  double y_split_      = 0.;
};

struct selection_t
{
  std::string pane_;
  std::string active_cell_;
  std::string sqref_;
};

/**
 * @brief Options for data validation.
 *
 * Options for data validation.
 *
 * @todo  Values can be mutualised (`std::variant`).
 */
struct data_validation_t
{
  /**
   * Set the validation type.
   *
   * Should be a `validation_types_t` value.
   */
  validation_types_t validate_ = validation_types_t::NONE;

  /**
   * Set the validation criteria type to select the data.
   *
   * Should be a `validation_criteria_t` value.
   */
  validation_criteria_t criteria_ = validation_criteria_t::NONE;

  /**
   * Controls whether a data validation is not applied to blank data in the
   * cell.
   *
   * Should be a `validation_boolean_t` value.
   *
   * It is on by default.
   */
  validation_boolean_t ignore_blank_ = validation_boolean_t::DEFAULT;

  /**
   * This parameter is used to toggle on and off the 'Show input message
   * when cell is selected' option in the Excel data validation dialog. When
   * the option is off an input message is not displayed even if it has been
   * set using input_message.
   *
   * Should be a `validation_boolean_t` value. It is on by default.
   */
  validation_boolean_t show_input_ = validation_boolean_t::DEFAULT;

  /**
   * This parameter is used to toggle on and off the 'Show error alert
   * after invalid data is entered' option in the Excel data validation
   * dialog. When the option is off an error message is not displayed even
   * if it has been set using error_message.
   *
   * Should be a `validation_boolean_t` value. It is on by default.
   */
  validation_boolean_t show_error_ = validation_boolean_t::DEFAULT;

  /**
   * This parameter is used to specify the type of error dialog that is
   * displayed.
   *
   * Should be a `validation_error_types_t` value.
   */
  validation_error_types_t error_type_ = validation_error_types_t::STOP;

  /**
   * This parameter is used to toggle on and off the 'In-cell dropdown'
   * option in the Excel data validation dialog. When the option is on a
   * dropdown list will be shown for list validations.
   *
   * Should be a `validation_boolean_t` value. It is on by default.
   */
  validation_boolean_t dropdown_ = validation_boolean_t::DEFAULT;

  /**
   * This parameter is used to set the limiting value to which the criteria
   * is applied using a whole or decimal number.
   */
  double value_number_ = 0.;

  /**
   * This parameter is used to set the limiting value to which the criteria
   * is applied using a cell reference.
   *
   * It is valid for any of the `_FORMULA` validation types.
   */
  std::string value_formula_;

  /**
   * This parameter is used to set a list of strings for a dropdown list.
   *
   * The `%value_formula_` parameter can also be used to specify a list from
   * an Excel cell range.
   *
   * @note The string list is restricted by Excel to 255 characters,
   * including comma separators.
   */
  std::vector<std::string> value_list_;

  /**
   * This parameter is used to set the limiting value to which the date or
   * time criteria is applied.
   *
   * @todo Manage internally with a port of lxw_datetime, some for other datetime.
   */
  std::chrono::system_clock::time_point value_datetime_;

  /**
   * This parameter is the same as `%value_number_` but for the minimum value
   * when a `BETWEEN` criteria is used.
   */
  double minimum_number_ = 0.;

  /**
   * This parameter is the same as `%value_formula_` but for the minimum value
   * when a `BETWEEN` criteria is used.
   */
  std::string minimum_formula_;

  /**
   * This parameter is the same as `%value_datetime_` but for the minimum value
   * when a `BETWEEN` criteria is used.
   */
  std::chrono::system_clock::time_point minimum_datetime_;

  /**
   * This parameter is the same as `%value_number_` but for the maximum value
   * when a `BETWEEN` criteria is used.
   */
  double maximum_number_ = 0.;

  /**
   * This parameter is the same as `%value_formula_` but for the maximum value
   * when a `BETWEEN` criteria is used.
   */
  std::string maximum_formula_;

  /**
   * This parameter is the same as `%value_datetime_` but for the maximum value
   * when a `BETWEEN` criteria is used.
   */
  std::chrono::system_clock::time_point maximum_datetime_;

  /**
   * The `%input_title_` parameter is used to set the title of the input message
   * that is displayed when a cell is entered.
   *
   * It has no default value and is only displayed if the input message
   * is displayed. The maximum title length is 32 characters.
   *
   * @see The `%input_message_` parameter below.
   */
  std::string input_title_;

  /**
   * The `%input_message parameter_` is used to set the input message that is
   * displayed when a cell is entered. It has no default value.
   *
   * The message can be split over several lines using newlines.
   * The maximum message length is 255 characters.
   */
  std::string input_message_;

  /**
   * The `%error_title_` parameter is used to set the title of the error message
   * that is displayed when the data validation criteria is not met.
   *
   * The default error title is 'Microsoft Excel'. The maximum title length is
   * 32 characters.
   */
  std::string error_title_;

  /**
   * The `%error_message_` parameter is used to set the error message that is
   * displayed when a cell is entered.
   *
   * The default error message is "The value you entered is not valid.
   * A user has restricted values that can be entered into the cell".
   * The message can be split over several lines using newlines. The maximum
   * message length is 255 characters.
   */
  std::string error_message_;
};

// A copy of `data_validation_t` which is used internally and which contains
// some additional fields.
struct data_val_obj_t
{
  validation_types_t validate_         = validation_types_t::NONE;
  validation_criteria_t criteria_      = validation_criteria_t::NONE;
  bool ignore_blank_                   = false;
  bool show_input_                     = false;
  bool show_error_                     = false;
  validation_error_types_t error_type_ = validation_error_types_t::STOP;
  bool dropdown_                       = false;
  double value_number_                 = 0.;
  std::string value_formula_;
  double minimum_number_ = 0.;
  std::string minimum_formula_;
  double maximum_number_ = 0.;
  std::string maximum_formula_;
  std::string input_title_;
  std::string input_message_;
  std::string error_title_;
  std::string error_message_;
  std::string sqref_;
};

/**
 * @brief Options for conditional formatting options.
 *
 * The fields/options in the the `%conditional_format_t` are used to define a
 * worksheet conditional format. It is used in conjunction with
 * `worksheet_t::conditional_format_cell()` and `worksheet_t::conditional_format_range()`.
 */
struct conditional_format_t
{
  /**
   * The type of conditional format such as `%conditional_format_types_t::CELL` or
   * `%conditional_format_types_t::DATA_BAR`.
   *
   * Should be a `conditional_format_types_t` value.
   */
  conditional_format_types_t type_ = conditional_format_types_t::NONE;

  /**
   * The criteria parameter is used to set the criteria by which the cell
   * data will be evaluated. For example in the expression `a > 5` the
   * criteria is `>` or, in `Xlsxwriter++` terms,
   * `%filter_criteria_t::GREATER_THAN`. The criteria that are
   * applicable depend on the conditional format type.
   *
   * The criteria options are defined in `conditional_criteria_t`.
   */
  conditional_criteria_t criteria_ = conditional_criteria_t::NONE;

  /**
   * The number value to which the condition refers. For example in the
   * expression `a > 5`, the value is 5.
   */
  double value_ = 0.;

  /**
   * The string value to which the condition refers, such as `"=A1"`. If a
   * value_string exists in the struct then the number value is
   * ignored.
   *
   * @note, if the condition refers to a text string then it must
   * be double quoted like this `"foo"`.
   */
  std::string value_string_;

  /**
   * The format field is used to specify the format_t format that will
   * be applied to the cell when the conditional formatting criterion is
   * met. The format_t is created using the `workbook_t::add_format()`
   * method in the same way as cell formats.
   *
   * @note In Excel, a conditional format is superimposed over the existing
   * cell format and not all cell format properties can be
   * modified. Properties that @b cannot be modified, in Excel, by a
   * conditional format are: font name, font size, superscript and
   * subscript, diagonal borders, all alignment properties and all
   * protection properties.
   */
  format_t* format_ = nullptr;

  /**
   * The minimum value used for Cell, Color Scale and Data Bar conditional
   * formats. For Cell types this is usually used with a "Between" style
   * criteria.
   */
  double min_value_ = 0.;

  /**
   * The minimum string value used for Cell, Color Scale and Data Bar conditional
   * formats. Usually used to set ranges like `=A1`.
   */
  std::string min_value_string_;

  /**
   * The rule used for the minimum condition in Color Scale and Data Bar
   * conditional formats. The rule types are defined in
   * `conditional_format_rule_types_t`.
   */
  conditional_format_rule_types_t min_rule_type_ = conditional_format_rule_types_t::NONE;

  /**
   * The color used for the minimum Color Scale conditional format.
   * See @ref working_with_colors.
   */
  color_t min_color_ = color_t::UNSET;

  /**
   * The middle value used for Color Scale and Data Bar conditional formats.
   */
  double mid_value_ = 0.;

  /**
   * The middle string value used for Color Scale and Data Bar conditional
   * formats. Usually used to set ranges like `=A1`.
   */
  std::string mid_value_string_;

  /**
   * The rule used for the middle condition in Color Scale and Data Bar
   * conditional formats. The rule types are defined in
   * `conditional_format_rule_types_t`.
   */
  conditional_format_rule_types_t mid_rule_type_ = conditional_format_rule_types_t::NONE;

  /**
   * The color used for the middle Color Scale conditional format.
   *
   * @see @ref working_with_colors.
   */
  color_t mid_color_ = color_t::UNSET;

  /**
   * The maximum value used for Cell, Color Scale and Data Bar conditional
   * formats. For Cell types this is usually used with a "Between" style
   * criteria.
   */
  double max_value_ = 0.;

  /**
   * The maximum string value used for Cell, Color Scale and Data Bar conditional
   * formats. Usually used to set ranges like `=A1`.
   */
  std::string max_value_string_;

  /**
   * The rule used for the maximum condition in Color Scale and Data Bar
   * conditional formats. The rule types are defined in
   * `conditional_format_rule_types_t`.
   */
  conditional_format_rule_types_t max_rule_type_ = conditional_format_rule_types_t::NONE;

  /**
   * The color used for the maximum Color Scale conditional format.
   *
   * @see @ref working_with_colors.
   */
  color_t max_color_ = color_t::UNSET;

  /**
   * The `%bar_color_` field sets the fill color for data bars.
   *
   * @see @ref working_with_colors.
   */
  color_t bar_color_ = color_t::UNSET;

  /**
   * The `%bar_only_` field sets The bar_only field displays a bar data but
   * not the data in the cells.
   */
  bool bar_only_ = false;

  /**
   * In Excel 2010 additional data bar properties were added such as solid
   * (non-gradient) bars and control over how negative values are
   * displayed. These properties can shown below.
   *
   * The `%data_bar_2010_` field sets Excel 2010 style data bars even when
   * Excel 2010 specific properties aren't used.
   */
  bool data_bar_2010_ = false;

  /**
   * The `%bar_solid_` field turns on a solid (non-gradient) fill for data
   * bars.
   *
   * @note Excel 2010 and higher only.
   */
  bool bar_solid_ = false;

  /**
   * The `%bar_negative_color_` field sets the color fill for the negative
   * portion of a data bar.
   *
   * @see @ref working_with_colors.
   *
   * @note Excel 2010 and higher only.
   */
  color_t bar_negative_color_ = color_t::UNSET;

  /**
   * The `%bar_border_color_` field sets the color for the border line of a
   * data bar.
   *
   * @see @ref working_with_colors.
   *
   * @note Excel 2010 and higher only.
   */
  color_t bar_border_color_ = color_t::UNSET;

  /**
   * The `%bar_negative_border_color_` field sets the color for the border of
   * the negative portion of a data bar.
   *
   * @see @ref working_with_colors.
   *
   * @note Excel 2010 and higher only.
   */
  color_t bar_negative_border_color_ = color_t::UNSET;

  /**
   * The `%bar_negative_color_same_` field sets the fill color for the negative
   * portion of a data bar to be the same as the fill color for the
   * positive portion of the data bar.
   *
   * @note Excel 2010 and higher only.
   */
  bool bar_negative_color_same_ = false;

  /**
   * The `%bar_negative_border_color_same_` field sets the border color for the
   * negative portion of a data bar to be the same as the border color for
   * the positive portion of the data bar.
   *
   * @note Excel 2010 and higher only.
   */
  bool bar_negative_border_color_same_ = false;

  /**
   * The `%bar_no_border_` field turns off the border for data bars.
   *
   * @note Excel 2010 and higher only.
   */
  bool bar_no_border_ = false;

  /**
   * The `%bar_direction_` field sets the direction for data bars. This
   * property can be either left for left-to-right or right for
   * right-to-left. If the property isn't set then Excel will adjust the
   * position automatically based on the context. Should be a
   * `conditional_format_bar_direction_t` value.
   *
   * @note Excel 2010 and higher only.
   */
  conditional_format_bar_direction_t bar_direction_ = conditional_format_bar_direction_t::CONTEXT;

  /**
   * The `%bar_axis_position_` field sets the position within the cells for the
   * axis that is shown in data bars when there are negative values to
   * display. The property can be either middle or none. If the property
   * isn't set then Excel will position the axis based on the range of
   * positive and negative values. Should be a
   * `conditional_bar_axis_position_t` value.
   *
   * @note Excel 2010 and higher only.
   */
  conditional_bar_axis_position_t bar_axis_position_ = conditional_bar_axis_position_t::AUTOMATIC;

  /**
   * The `%bar_axis_color_` field sets the color for the axis that is shown
   * in data bars when there are negative values to display.
   *
   * @see @ref working_with_colors.
   *
   * @note Excel 2010 and higher only.
   */
  color_t bar_axis_color_ = color_t::UNSET;

  /**
   * The Icons Sets style is specified by the `%icon_style_` parameter. Should
   * be a `conditional_icon_types_t`.
   */
  conditional_icon_types_t icon_style_ = conditional_icon_types_t::THREE_ARROWS_COLORED;

  /**
   * The order of Icon Sets icons can be reversed by setting `%reverse_icons_`.
   */
  bool reverse_icons_ = false;

  /**
   * The icons can be displayed without the cell value by settings the
   *  `%icons_only_` parameter.
   */
  bool icons_only_ = false;

  /**
   * The multi_range field is used to extend a conditional format over
   * non-contiguous ranges.
   *
   * It is possible to apply the conditional format to different cell
   * ranges in a worksheet using multiple calls to `worksheet_t::conditional_format_cell()`.
   * However, as a minor optimization it is also possible in Excel to apply the same conditional format to
   * different non-contiguous cell ranges with `worksheet_t::conditional_format_range()`.
   *
   *  This is replicated in `worksheet_t::conditional_format_range()` using the
   *  multi_range option. The range must contain the primary range for the
   *  conditional format and any others separated by spaces. For example
   *  `"A1 C1:C5 E2 G1:G100"`.
   */
  std::string multi_range_;

  /**
   * The `%stop_if_true_` parameter can be used to set the "stop if true"
   * feature of a conditional formatting rule when more than one rule is
   * applied to a cell or a range of cells. When this parameter is set then
   * subsequent rules are not evaluated if the current rule is true.
   */
  bool stop_if_true_ = false;
};

struct cond_format_obj_t
{
  conditional_format_types_t type_ = conditional_format_types_t::NONE;
  conditional_criteria_t criteria_ = conditional_criteria_t::NONE;
  double min_value_                = 0.;
  std::string min_value_string_;
  conditional_format_rule_types_t min_rule_type_ = conditional_format_rule_types_t::NONE;
  color_t min_color_                             = color_t::UNSET;
  double mid_value_                              = 0.;
  std::string mid_value_string_;
  // TODO ?    uint8_t mid_value_type;
  conditional_format_rule_types_t mid_rule_type_ = conditional_format_rule_types_t::NONE;
  color_t mid_color_                             = color_t::UNSET;
  double max_value_                              = 0.;
  std::string max_value_string_;
  // TODO ?     uint8_t max_value_type;
  conditional_format_rule_types_t max_rule_type_     = conditional_format_rule_types_t::NONE;
  color_t max_color_                                 = color_t::UNSET;
  bool data_bar_2010_                                = false;
  bool auto_min_                                     = false;
  bool auto_max_                                     = false;
  bool bar_only_                                     = false;
  bool bar_solid_                                    = false;
  bool bar_negative_color_same_                      = false;
  bool bar_negative_border_color_same_               = false;
  bool bar_no_border_                                = false;
  conditional_format_bar_direction_t bar_direction_  = conditional_format_bar_direction_t::CONTEXT;
  conditional_bar_axis_position_t bar_axis_position_ = conditional_bar_axis_position_t::AUTOMATIC;
  color_t bar_color_                                 = color_t::UNSET;
  color_t bar_negative_color_                        = color_t::UNSET;
  color_t bar_border_color_                          = color_t::UNSET;
  color_t bar_negative_border_color_                 = color_t::UNSET;
  color_t bar_axis_color_                            = color_t::UNSET;
  conditional_icon_types_t icon_style_               = conditional_icon_types_t::THREE_ARROWS_COLORED;
  bool reverse_icons_                                = false;
  bool icons_only_                                   = false;
  bool stop_if_true_                                 = false;
  bool has_max_                                      = false;
  std::string type_string_;
  std::string guid_;
  // PROPERTY_UNSET should not be part of format_t
  int32_t dxf_index_     = format_t::PROPERTY_UNSET;
  uint32_t dxf_priority_ = 0;
  std::string first_cell_;
  std::string sqref_;
};

/**
 * @brief Table columns options.
 *
 * Structure to set the options of a table column added with
 * `worksheet_t::add_table()`.
 *
 * @see @ref ww_tables_columns.
 */
struct table_column_t
{
  /**
   * Set the header name/caption for the column. If empty the header defaults
   * to `Column 1`, `Column 2`, etc.
   */
  std::string header_;

  /**
   * Set the formula for the column.
   */
  std::string formula_;

  /**
   * Set the string description for the column total.
   */
  std::string total_string_;

  /**
   * Set the function for the column total.
   */
  table_total_functions_t total_function_ = table_total_functions_t::NONE;

  /**
   * Set the format for the column header.
   */
  format_t* header_format_;

  /**
   * Set the format for the data rows in the column.
   */
  format_t* format_;

  /**
   * Set the formula value for the column total (not generally required).
   */
  double total_value_ = 0.;
};

/**
 * @brief Worksheet table options.
 *
 * Options used to define worksheet tables.
 *
 * @see @ref working_with_tables for more information.
 */
struct table_options_t
{
  // clang-format off
  /**
   * The `name` parameter is used to set the name of the table. This
   * parameter is optional and by default tables are named `Table1`,
   * `Table2`, etc. in the worksheet order that they are added.
   *
   * @code
   *  xwpp::table_options_t options{.name_ = "Sales"};
   *
   *  worksheet.add_table(RANGE("B3:G8"), options);
   * @endcode
   *
   * If you override the table name you must ensure that it doesn't clash
   * with an existing table name and that it follows Excel's requirements
   * for table names, see the Microsoft Office documentation on
   * [Naming an Excel Table](https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114).
   */
  // clang-format on
  std::string name_;

  /**
   * The `%no_header_row_` parameter can be used to turn off the header row in
   * the table. It is on by default:
   *
   * @code
   *  xwpp::table_options_t options{.no_header_row_ = true};
   *
   *  worksheet.add_table(RANGE("B4:F7"), options);
   * @endcode
   *
   * @image html tables4.png
   *
   * Without this option the header row will contain default captions such
   * as `Column 1`, `Column 2`, etc. These captions can be overridden
   * using the `%columns_` parameter shown below.
   */
  bool no_header_row_ = false;

  /**
   * The `%no_autofilter_` parameter can be used to turn off the autofilter in
   * the header row. It is on by default:
   *
   * @code
   *  xwpp::table_options_t options{.no_autofilter_ = true};
   *
   *  worksheet.add_table(RANGE("B3:F7"), options);
   * @endcode
   *
   * @image html tables3.png
   *
   * The autofilter is only shown if the `%no_header_row_` parameter is off
   * (the default). Filter conditions within the table are not supported.
   */
  bool no_autofilter_ = false;

  /**
   * The `%no_banded_rows_` parameter can be used to turn off the rows of
   * alternating color in the table. It is on by default:
   *
   * @code
   *  xwpp::table_options_t options{.no_banded_rows_ = true};
   *
   *  worksheet.add_table(RANGE("B3:F7"), options);
   * @endcode
   *
   * @image html tables6.png
   */
  bool no_banded_rows_ = false;

  /**
   * The `%banded_columns_` parameter can be used to used to create columns of
   * alternating color in the table. It is off by default:
   *
   * @code
   *  xwpp::table_options_t options{.banded_columns_ = true};
   *
   *  worksheet.add_table(RANGE("B3:F7"), options);
   * @endcode
   *
   * The banded columns formatting is shown in the image in the previous
   * section above.
   */
  bool banded_columns_ = false;

  /**
   * The `first_column` parameter can be used to highlight the first column
   * of the table. The type of highlighting will depend on the `style_type`
   * of the table. It may be bold text or a different color. It is off by
   * default:
   *
   * @code
   *  xwpp::table_options_t options{
   *    .first_column_ = true,
   *    .last_column_ = true,
   *  };
   *
   *  worksheet.add_table(RANGE("B3:F7"), options);
   * @endcode
   *
   * @image html tables5.png
   */
  bool first_column_ = false;

  /**
   * The `%last_column_` parameter can be used to highlight the last column of
   * the table. The type of highlighting will depend on the `style` of the
   * table. It may be bold text or a different color. It is off by default:
   *
   * @code
   *  xwpp::table_options_t options{
   *    .first_column_ = true,
   *    .last_column_ = true,
   *  };
   *
   *  worksheet.add_table(RANGE("B3:F7"), options);
   * @endcode
   *
   * The `%last_column_` formatting is shown in the image in the previous
   * section above.
   */
  bool last_column_ = false;

  /**
   * The `%style_type_` parameter can be used to set the style of the table,
   * in conjunction with the `%style_type_number_` parameter:
   *
   * @code
   *  xwpp::table_options_t options{
   *    .style_type_        = xwpp::table_style_type_t::LIGHT,
   *    .style_type_number_ = 11,
   *  };
   *
   *  worksheet.add_table(RANGE("B3:G8"), options);
   * @endcode
   *
   * @image html tables11.png
   *
   * There are three types of table style in Excel: Light, Medium and Dark
   * which are represented using the #table_style_type_t enum values:
   *
   * - `%table_style_type_t::LIGHT`
   * - `%table_style_type_t::MEDIUM`
   * - `%table_style_type_t::DARK`
   *
   * Within those ranges there are between 11 and 28 other style types which
   * can be set with `%style_type_number_` (depending on the style type).
   * Check Excel to find the style that you want. The dialog with the
   * options laid out in numeric order are shown below:
   *
   * @image html tables14.png
   *
   * The default table style in Excel is 'Table Style Medium 9' (highlighted
   * with a green border in the image above), which is set by default in
   * Xlsxwriter++ as:
   *
   * @code
   *  xwpp::table_options_t options{
   *    .style_type_        = xwpp::table_style_type_t::MEDIUM,
   *    .style_type_number_ = 9,
   *  };
   * @endcode
   *
   * You can also turn the table style off by setting it to Light 0:
   *
   * @code
   *  xwpp::table_options_t options{
   *    .style_type_        = xwpp::table_style_type_t::LIGHT,
   *    .style_type_number_ = 0,
   *  };
   * @endcode
   *
   * @image html tables13.png
   */
  table_style_type_t style_type_ = table_style_type_t::DEFAULT;

  /**
   * The `%style_type_number_` parameter is used with `%style_type_` to set the
   * style of a worksheet table.
   */
  uint8_t style_type_number_ = 0;

  /**
   * The `%total_row_` parameter can be used to turn on the total row in the
   * last row of a table. It is distinguished from the other rows by a
   * different formatting and also with dropdown `SUBTOTAL` functions:
   *
   * @code
   *  xwpp::table_options_t options{.total_row_ = true};
   *
   *  worksheet.add_table(RANGE("B3:G8"), options);
   * @endcode
   *
   * @image html tables9.png
   *
   * The default total row doesn't have any captions or functions. These
   * must by specified via the `%columns_` parameter below.
   */
  bool total_row_ = false;

  /**
   * The `%columns_` parameter can be used to set properties for columns
   * within the table.
   *
   * @see @ref ww_tables_columns for a detailed explanation.
   */
  std::vector<table_column_t> columns_;
};

struct table_obj_t
{
  std::string name_;
  std::string total_string_;
  std::vector<table_column_t> columns_;
  bool banded_columns_           = false;
  bool first_column_             = false;
  bool last_column_              = false;
  bool no_autofilter_            = false;
  bool no_banded_rows_           = false;
  bool no_header_row_            = false;
  table_style_type_t style_type_ = table_style_type_t::DEFAULT;
  uint8_t style_type_number_     = 0;
  bool total_row_                = false;
  row_num_t first_row_           = 0;
  col_num_t first_col_           = 0;
  row_num_t last_row_            = 0;
  col_num_t last_col_            = 0;
  col_num_t num_cols_            = 0; // TODO is it useful?
  uint32_t id_                   = 0;
  std::string sqref_;
  std::string filter_sqref_;
};

/**
 * @brief Options for autofilter rules.
 *
 * Options to define an autofilter rule.
 */
struct filter_rule_t
{
  /**
   * The filter_criteria_t to define the rule.
   */
  filter_criteria_t criteria_ = filter_criteria_t::NONE;

  /**
   * String value to which the criteria applies.
   */
  std::string value_string_;

  /**
   * Numeric value to which the criteria applies (if `%value_string_` isn't used).
   */
  double value_ = 0;
};

struct filter_rule_obj_t
{
  filter_type_t type_          = filter_type_t::NONE;
  bool is_custom_              = false;
  bool has_blanks_             = false;
  col_num_t col_num_           = 0;
  filter_criteria_t criteria1_ = filter_criteria_t::NONE;
  filter_criteria_t criteria2_ = filter_criteria_t::NONE;
  double value1_               = 0.;
  double value2_               = 0.;
  std::string value1_string_;
  std::string value2_string_;
  std::vector<std::string> list_;
};

/**
 * @brief Options for inserted images.
 *
 * Options for modifying images inserted via `worksheet_t::insert_image()` and
 * `worksheet_t::embed_image()`.
 */
struct image_options_t
{
  /**
   * Offset from the left of the cell in pixels.
   */
  int32_t x_offset_ = 0;

  /**
   * Offset from the top of the cell in pixels.
   */
  int32_t y_offset_ = 0;

  /**
   *  X scale of the image as a decimal.
   */
  double x_scale_ = 1.;

  /**
   *  Y scale of the image as a decimal.
   */
  double y_scale_ = 1.;

  /**
   * Object position - use one of the values of `object_position_t`.
   *
   * @see @ref working_with_object_positioning.
   */
  object_position_t object_position_ = object_position_t::DEFAULT;

  /**
   * Optional description or "Alt text" for the image. This field can be
   * used to provide a text description of the image to help
   * accessibility.
   *
   * Defaults to the image filename as in Excel.
   * Set to empty string to ignore the description field.
   */
  std::optional<std::string> description_;

  /**
   * Optional parameter to help accessibility. It is used to mark the image
   * as decorative, and thus uninformative, for automated screen
   * readers. As in Excel, if this parameter is in use the `%description_`
   * field isn't written.
   */
  bool decorative_ = false;

  /**
   * Add an optional hyperlink to the image. Follows the same URL rules
   * and types as `worksheet_t::write_url()`.
   */
  std::string url_;

  /**
   * Add an optional mouseover tip for a hyperlink to the image.
   */
  std::string tip_;

  /**
   * Add an optional format to the cell. Only used with `worksheet_t::embed_image()`.
   */
  format_t* cell_format_;
};

/**
 * @brief Options for inserted charts.
 *
 * Options for modifying charts inserted via `worksheet_t::insert_chart()`.
 *
 */
struct chart_options_t
{
  /**
   * Offset from the left of the cell in pixels.
   */
  int32_t x_offset_ = 0;

  /**
   * Offset from the top of the cell in pixels.
   */
  int32_t y_offset_ = 0;

  /**
   * X scale of the chart as a decimal.
   */
  double x_scale_ = 0.;

  /**
   * Y scale of the chart as a decimal.
   */
  double y_scale_ = 0.;

  /**
   * Object position - use one of the values of `object_position_t`.
   *
   * @see @ref working_with_object_positioning.
   */
  object_position_t object_position_ = object_position_t::DEFAULT;

  /**
   * Optional description or "Alt text" for the chart. This field can be
   * used to provide a text description of the chart to help
   * accessibility.
   *
   * Defaults to the image filename as in Excel. Set to empty string
   * to ignore the description field.
   */
  std::string description_;

  /**
   * Optional parameter to help accessibility. It is used to mark the chart
   * as decorative, and thus uninformative, for automated screen
   * readers. As in Excel, if this parameter is in use the `%description_`
   * field isn't written.
   */
  bool decorative_ = false;
};

// Internal struct to represent `image_options_t` and `chart_options_t`
// values as well as internal metadata.
struct object_properties_t
{
  int32_t x_offset_ = 0;
  int32_t y_offset_ = 0;
  double x_scale_   = 1.;
  double y_scale_   = 1.;
  row_num_t row_;
  col_num_t col_;
  std::string filename_;
  std::string description_;
  std::string url_;
  std::string tip_;
  object_position_t object_position_ = object_position_t::DEFAULT;
  image_types_t image_type_;
  std::vector<unsigned char> image_buffer_;
  double width_  = 0.;
  double height_ = 0.;
  std::string extension_;
  double x_dpi_;
  double y_dpi_;
  chart_t* chart_     = nullptr;
  bool is_duplicate_  = false;
  bool is_background_ = false;
  std::string md5_;
  std::string image_position_;
  bool decorative_  = false;
  format_t* format_ = nullptr;
};

/**
 * @brief Options for inserted comments.
 *
 * Options for modifying comments inserted via `worksheet_t::write_comment()`.
 */
struct comment_options_t
{
  /**
   * This option is used to make a cell comment visible when the worksheet
   * is opened. The default behavior in Excel is that comments are
   * initially hidden. However, it is also possible in Excel to make
   * individual comments or all comments visible.  You can make all
   * comments in the worksheet visible using the
   * `worksheet_t::show_comments()` function. Defaults to
   * `comment_display_t::DEFAULT`.
   *
   * @see @ref ww_comments_visible.
   */
  comment_display_t visible_ = comment_display_t::DEFAULT;

  /**
   * This option is used to indicate the author of the cell comment. Excel
   * displays the author in the status bar at the bottom of the
   * worksheet. The default author for all cell comments in a worksheet can
   * be set using the `worksheet_t::set_comments_author()` function. Set to
   * empty string if not required.
   *
   * @see @ref ww_comments_author.
   */
  std::string author_;

  /**
   * This option is used to set the width of the cell comment box
   * explicitly in pixels. The default width is 128 pixels.
   *
   * @see @ref ww_comments_width.
   */
  uint16_t width_ = 128;

  /**
   * This option is used to set the height of the cell comment box
   * explicitly in pixels. The default height is 74 pixels.
   *
   * @see @ref ww_comments_height.
   */
  uint16_t height_ = 74;

  /**
   * X scale of the comment as a decimal.
   *
   * @see @ref ww_comments_x_scale.
   */
  double x_scale_ = 1.;

  /**
   * Y scale of the comment as a decimal.
   *
   * @see @ref ww_comments_y_scale.
   */
  double y_scale_ = 1.;

  /**
   * This option is used to set the background color of cell comment
   * box. The color should be an RGB integer value.
   *
   * @see @ref working_with_colors.
   * @see @ref ww_comments_color.
   */
  color_t color_ = color_t::UNSET;

  /**
   * This option is used to set the font for the comment. The default font
   * is 'Tahoma'.
   *
   * @see @ref ww_comments_font_name.
   */
  std::string font_name_ = "Tahoma";

  /**
   * This option is used to set the font size for the comment. The default
   * is 8.
   *
   * @see @ref ww_comments_font_size.
   */
  double font_size_ = 8.;

  /**
   * This option is used to set the font family number for the comment.
   * Not required very often. Set to 0.
   */
  uint8_t font_family_ = 0;

  /**
   * This option is used to set the row in which the comment will
   * appear. By default Excel displays comments one cell to the right and
   * one cell above the cell to which the comment relates. The `%start_row_`
   * and `%start_col_` options should both be set to 0 if not used.
   *
   * @see @ref ww_comments_start_row.
   */
  row_num_t start_row_ = 0;

  /**
   * This option is used to set the column in which the comment will
   * appear. See the `%start_row_` option for more information.
   *
   * @see @ref ww_comments_start_col.
   */
  col_num_t start_col_ = 0;

  /**
   * Offset from the left of the cell in pixels.
   *
   * @see @ref ww_comments_x_offset.
   */
  int32_t x_offset_ = 0;

  /**
   * Offset from the top of the cell in pixels.
   *
   * @see @ref ww_comments_y_offset.
   */
  int32_t y_offset_ = 0;
};

/**
 * @brief Options for inserted buttons.
 *
 * Options for modifying buttons inserted via `worksheet_t::insert_button()`.
 *
 */
struct button_options_t
{
  /**
   * Sets the caption on the button. The default is "Button n" where n is
   * the current number of buttons in the worksheet, including this
   * button.
   */
  std::string caption_;

  /**
   * Name of the macro to run when the button is pressed. The macro must be
   * included with `workbook_t::add_vba_project()`.
   */
  std::string macro_;

  /**
   * Optional description or "Alt text" for the button. This field can be
   * used to provide a text description of the button to help
   * accessibility. Set to empty string to ignore the description field.
   */
  std::string description_;

  /**
   * This option is used to set the width of the cell button box
   * explicitly in pixels. The default width is 64 pixels.
   */
  uint16_t width_ = 0;

  /**
   * This option is used to set the height of the cell button box
   * explicitly in pixels. The default height is 20 pixels.
   */
  uint16_t height_ = 0;

  /**
   * X scale of the button as a decimal.
   */
  double x_scale_ = 0.;

  /**
   * Y scale of the button as a decimal.
   */
  double y_scale_ = 0.;

  /**
   * Offset from the left of the cell in pixels.
   */
  int32_t x_offset_ = 0;

  /**
   * Offset from the top of the cell in pixels.
   */
  int32_t y_offset_ = 0;
};

/**
 * @brief Options for headers and footers.
 *
 * Optional parameters used in the `worksheet_t::set_header()` and
 * `worksheet_t::set_footer()` functions.
 */
struct header_footer_options_t
{
  /**
   * Header or footer margin in inches. Excel default is 0.3. Must by
   * larger than 0.0.
   *
   * @see `worksheet_t::set_header()`.
   */
  double margin_;

  /**
   * The left header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&L` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_left_;

  /**
   * The center header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&C` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_center_;

  /**
   * The right header image filename, with path if required. This should
   * have a corresponding `&G` placeholder in the `&R` section of
   * the header/footer string.
   *
   * @see `worksheet_t::set_header()`.
   */
  std::string image_right_;
};

/**
 * @brief Worksheet protection options.
 *
 * Worksheet protection options.
 */
struct protection_t
{
  /**
   * Turn off selection of locked cells. This in on in Excel by default.
   */
  bool no_select_locked_cells_ = false;

  /**
   * Turn off selection of unlocked cells. This in on in Excel by default.
   */
  bool no_select_unlocked_cells_ = false;

  /**
   * Prevent formatting of cells.
   */
  bool format_cells_ = false;

  /**
   * Prevent formatting of columns.
   */
  bool format_columns_ = false;

  /**
   * Prevent formatting of rows.
   */
  bool format_rows_ = false;

  /**
   * Prevent insertion of columns.
   */
  bool insert_columns_ = false;

  /**
   * Prevent insertion of rows.
   */
  bool insert_rows_ = false;

  /**
   * Prevent insertion of hyperlinks.
   */
  bool insert_hyperlinks_ = false;

  /**
   * Prevent deletion of columns.
   */
  bool delete_columns_ = false;

  /**
   * Prevent deletion of rows.
   */
  bool delete_rows_ = false;

  /**
   * Prevent sorting data.
   */
  bool sort_ = false;

  /**
   * Prevent filtering data.
   */
  bool autofilter_ = false;

  /**
   * Prevent insertion of pivot tables.
   */
  bool pivot_tables_ = false;

  /**
   * Protect scenarios.
   */
  bool scenarios_ = false;

  /**
   * Protect drawing objects. Worksheets only.
   *
   * @todo Not clear, why there is two options `%objects_` and `%no_objects_`.
   * @todo And `%objects_` seems to be wrong name as XML object option is set to 1 iff objects_ is false.
   */
  bool objects_ = false;

  /**
   * Turn off chartsheet content protection.
   */
  bool no_content_ = false;

  /**
   * Turn off chartsheet objects.
   */
  bool no_objects_ = false;
};

// Internal struct to copy protection options and internal metadata.
struct protection_obj_t
{
  bool no_select_locked_cells_   = false;
  bool no_select_unlocked_cells_ = false;
  bool format_cells_             = false;
  bool format_columns_           = false;
  bool format_rows_              = false;
  bool insert_columns_           = false;
  bool insert_rows_              = false;
  bool insert_hyperlinks_        = false;
  bool delete_columns_           = false;
  bool delete_rows_              = false;
  bool sort_                     = false;
  bool autofilter_               = false;
  bool pivot_tables_             = false;
  bool scenarios_                = false;
  bool objects_                  = false;
  bool no_content_               = false;
  bool no_objects_               = false;
  bool no_sheet_                 = false;
  bool is_configured_            = false;
  std::string hash_;
};

/**
 * @brief Struct to represent a rich string format/string pair.
 *
 * Arrays of this struct are used to define "rich" multi-format strings that
 * are passed to `worksheet_t::write_rich_string()`. Each struct represents a
 * fragment of the rich multi-format string with a format_t to define the
 * format for the string part. If the string fragment is unformatted then
 * `nullptr` can be used for the format.
 */
struct rich_string_tuple_t
{
  /**
   * The format for a string fragment in a rich string. `nullptr` if the string
   * isn't formatted.
   */
  format_t* format_ = nullptr;

  /**
   * The string fragment.
   */
  std::string str_;
};

// Worksheet initialization data.
struct worksheet_init_data_t
{
  uint16_t index_;
  bool hidden_            = false;
  uint16_t* active_sheet_ = nullptr;
  uint16_t* first_sheet_  = nullptr;
  shared_strings_t* sst_;
  std::string name_;
  std::string quoted_name_;
  format_t* default_url_format_;
  uint16_t max_url_length_;
  bool use_1904_epoch_ = false;
};

/**
 * @brief Representation of an Excel worksheet.
 *
 * `%worksheet_t` represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * A Worksheet object isn't created directly. Instead a worksheet is
 * created by calling the `workbook_t::add_worksheet()` function from a
 * Workbook object:
 *
 * @code
 *  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
 * @endcode

 * @todo Add API with different for:
 *  - integer (int, short, long, unsigned, uint16_t, ...).
 *  - date (chrono, time_t, string that contains date, ...).
 *  - ...
 * @todo Add API with col and row names instead of number.
 * @todo Replace few overload by default-valued parameters (in particular `nullptr` for `format`)
 */
class worksheet_t
{
public:
  // TODO Constructor should not be public but only used by `workbook_t`.
  worksheet_t();
  worksheet_t(const worksheet_init_data_t& init_data, std::function<int32_t(format_t*)> get_xf_index,
              std::function<int32_t(format_t*)> get_dxf_index);

  /**
   * @brief Set the properties for one or more columns of cells with options.
   *
   * @param first_col The zero indexed first column.
   * @param last_col  The zero indexed last column.
   * @param width     The width of the column(s).
   * @param format    A pointer to a Format instance or `nullptr`.
   * @param options   Optional row parameters: hidden, level, collapsed.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_column()` function can be used to change the default
   * properties of a single column or a range of columns:
   *
   * @code
   *  // Width of columns B:D set to 30.
   *  worksheet.set_column(1, 3, 30);
   * @endcode
   *
   * If `%set_column()` is applied to a single column the value of
   * `%first_col` and `%last_col` should be the same:
   *
   * @code
   *  // Width of column B set to 30.
   *  worksheet.set_column(1, 1, 30);
   * @endcode
   *
   * It is also possible, and generally clearer, to specify a column range using
   * the form of `COLS()` macro:
   *
   * @code
   *  worksheet.set_column(4, 4, 20);
   *  worksheet.set_column(5, 8, 30);
   *
   *  // Same as the examples above but clearer.
   *  worksheet.set_column(COLS("E:E"), 20);
   *  worksheet.set_column(COLS("F:H"), 30);
   * @endcode
   *
   * The `%width` parameter sets the column width in the same units used by Excel
   * which is: the number of characters in the default font. The default width
   * is 8.43 in the default font of `Calibri 11`. The actual relationship between
   * a string width and a column width in Excel is complex. See the
   * [following explanation of column widths](https://support.microsoft.com/en-us/kb/214123)
   * from the Microsoft support documentation for more details. To set the width in pixels use the
   * `set_column_pixels()` function.
   *
   * There is no way to specify "AutoFit" for a column in the Excel file
   * format. This feature is only available at runtime from within Excel. It is
   * possible to simulate "AutoFit" in your application by tracking the maximum
   * width of the data in the column as your write it and then adjusting the
   * column width at the end.
   *
   * As usual the @ref format.h `format` parameter is optional. If you wish to
   * set the format without changing the width you can pass a default column
   * width of DEF_COL_WIDTH:
   *
   * @code
   *  xwpp::format_t* bold = workbook.add_format();
   *  bold->set_bold();
   *
   *  // Set the first column to bold.
   *  worksheet.set_column(0, 0, xwpp::DEF_COL_WIDTH, bold);
   * @endcode
   *
   * The `format` parameter will be applied to any cells in the column that
   * don't have a format. For example:
   *
   * @code
   *  // Column 1 has format1.
   *  worksheet.set_column(COLS("A:A"), 8.43, format1);
   *
   *  // Cell A1 in column 1 defaults to format1.
   *  worksheet.write_string(0, 0, "Hello");
   *
   *  // Cell A2 in column 1 keeps format2.
   *  worksheet.write_string(1, 0, "Hello", format2);
   * @endcode
   *
   * As in Excel a row format takes precedence over a default column format:
   *
   * @code
   *  // Row 1 has format1.
   *  worksheet.set_row(0, 15, format1);
   *
   *  // Col 1 has format2.
   *  worksheet.set_column(COLS("A:A"), 8.43, format2);
   *
   *  // Cell A1 defaults to format1, the row format.
   *  worksheet.write_string(0, 0, "Hello");
   *
   *  // Cell A2 keeps format2, the column format.
   *  worksheet.write_string(1, 0, "Hello");
   * @endcode
   *
   * The optional `options` parameter is a `row_col_options_t` struct. It has the
   * following members:
   *
   * - `%hidden_`
   * - `%level_`
   * - `%collapsed_`
   *
   * The `%hidden_` option is used to hide a column. This can be used, for
   * example, to hide intermediary steps in a complicated calculation:
   *
   * @code
   *  xwpp::row_col_options_t options = {.hidden_ = true, .level_ = 0, .collapsed_ = false};
   *
   *  worksheet.set_column(COLS("D:E"), xwpp::DEF_COL_WIDTH, nullptr, options);
   * @endcode
   *
   * @image html hide_row_col3.png
   *
   * The `%hidden_`, `%level_`,  and `%collapsed_` options can also be used to
   * create Outlines and Grouping. See @ref working_with_outlines.
   *
   * @code
   *  xwpp::row_col_options_t options = {.hidden_ = false, .level = 1, .collapsed = false};
   *
   *  worksheet.set_column(COLS("B:G"), 5, nullptr, &options1);
   * @endcode
   *
   * @image html outline8.png
   *
   * @todo Add API with only format or option (no width).
   * @todo Add API with options but no format.
   */
  void set_column(col_num_t first_col, col_num_t last_col, double width, const format_t* format,
                  const std::optional<row_col_options_t>& options);
  /// @overload
  void set_column(col_num_t first_col, col_num_t last_col, double width, const format_t* format);
  /// @overload
  void set_column(col_num_t first_col, col_num_t last_col, double width);

  /**
   * @brief Set the properties for a row of cells.
   *
   * @param row_num The zero indexed row number.
   * @param height  The row height.
   * @param format  A pointer to a Format instance or `nullptr`.
   * @param options Optional row parameters: hidden, level, collapsed.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_row()` function is used to change the default properties
   * of a row. The most common use for this function is to change the
   * height of a row:
   *
   * @code
   *  // Set the height of Row 1 to 20.
   *  worksheet.set_row(0, 20);
   * @endcode
   *
   * The height is specified in character units. To specify the height in pixels
   * use the `set_row_pixels()` function.
   *
   * The other common use for `%set_row()` is to set the a @ref
   * format.h "Format" for all cells in the row:
   *
   * @code
   *  xwpp::format_t* bold = workbook.add_format();
   *  bold->set_bold();
   *
   *  // Set the header row to bold.
   *  worksheet.set_row(0, 15, bold);
   * @endcode
   *
   * If you wish to set the format of a row without changing the height you can
   * pass the default row height of `DEF_ROW_HEIGHT`:
   *
   * @code
   *  worksheet.set_row(0, xwpp::DEF_ROW_HEIGHT, format);
   *  worksheet.set_row(0, 15, format); // Same as above.
   * @endcode
   *
   * The `format` parameter will be applied to any cells in the row that don't
   * have a format. As with Excel the row format is overridden by an explicit
   * cell format. For example:
   *
   * @code
   *  // Row 1 has format1.
   *  worksheet.set_row(0, 15, format1);
   *
   *  // Cell A1 in Row 1 defaults to format1.
   *  worksheet.write_string(0, 0, "Hello");
   *
   *  // Cell B1 in Row 1 keeps format2.
   *  worksheet.write_string(0, 1, "Hello", format2);
   * @endcode
   *
   * The `options` parameter is a `row_col_options_t` struct. It has the
   * following members:
   *
   * - `%hidden_`
   * - `%level_`
   * - `%collapsed_`
   *
   * The `%hidden_` option is used to hide a row. This can be used, for
   * example, to hide intermediary steps in a complicated calculation:
   *
   * @code
   *  xwpp::row_col_options_t options = {.hidden_ = true, .level_ = 0, .collapsed_ = false};
   *
   *  // Hide the fourth and fifth (zero indexed) rows.
   *  worksheet.set_row(3, 15, nullptr, options);
   *  worksheet.set_row(4, 15, nullptr, options);
   * @endcode
   *
   * @image html hide_row_col2.png
   *
   * The `%hidden_`, `%level_`,  and `%collapsed_` options can also be used to
   * create Outlines and Grouping. See @ref working_with_outlines.
   *
   * @code
   *  // The option structs with the outline level set.
   *  xwpp::row_col_options_t options1 = {.hidden_ = false, .level_ = 2, .collapsed_ = false};
   *  xwpp::row_col_options_t options2 = {.hidden_ = false, .level_ = 1, .collapsed_ = false};
   *
   *  // Set the row options with the outline level.
   *  worksheet.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options2);
   *  worksheet.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
   *  worksheet.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options2);
   * @endcode
   *
   * @image html outline1.png
   *
   * @todo Add API with only format or option (no height).
   * @todo Add API with options but no format.
   */
  void set_row(row_num_t row_num, double height, const format_t* format,
               const std::optional<row_col_options_t>& options);
  /// @overload
  void set_row(row_num_t row_num, double height, const format_t* format);
  /// @overload
  void set_row(row_num_t row_num, double height);

  /**
   * @brief Set the properties for one or more columns of cells, with the width
   *        in pixels.
   *
   * @param first_col The zero indexed first column.
   * @param last_col  The zero indexed last column.
   * @param pixels    The width of the column(s) in pixels.
   * @param format    A pointer to a Format instance or nullptr.
   * @param options   Optional row parameters: hidden, level, collapsed.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_column_pixels()` function is similar to
   * `set_column()` function except that the width can be set in
   * pixels:
   *
   * @code
   *  // Column width set to 75 pixels, the same as 10 character units.
   *  worksheet.set_column_pixels(5, 5, 75);
   * @endcode
   *
   * @image html set_column_pixels.png
   *
   * If you wish to set the format of a column without changing the width you can
   * pass the default column width in pixels: `DEF_COL_WIDTH_PIXELS`.
   *
   * @todo Add API with only format or options.
   */
  void set_column_pixels(col_num_t first_col, col_num_t last_col, uint32_t pixels, const format_t* format,
                         const std::optional<row_col_options_t>& options);
  /// @overload
  void set_column_pixels(col_num_t first_col, col_num_t last_col, uint32_t pixels);

  /**
   * @brief Set the properties for a row of cells, with the height in pixels.
   *
   * @param row_num The zero indexed row number.
   * @param pixels  The row height in pixels.
   * @param format  A pointer to a Format instance or `nullptr`.
   * @param options Optional row parameters: hidden, level, collapsed.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_row_pixels()` function is the same as the
   * `set_row()` function except that the height can be set in pixels.
   *
   * @code
   *  // Set the height of Row 1 to 20 pixels.
   *  worksheet.set_row_pixels(0, 20, NULL);
   * @endcode
   *
   * If you wish to set the format of a row without changing the height you can
   * pass the default row height in pixels: `DEF_ROW_HEIGHT_PIXELS`.
   */
  void set_row_pixels(row_num_t row_num, uint32_t pixels, const format_t* format,
                      const std::optional<row_col_options_t>& options);
  /// @overload
  void set_row_pixels(row_num_t row_num, uint32_t pixels, const format_t* format);
  /// @overload
  void set_row_pixels(row_num_t row_num, uint32_t pixels);

  /**
   * @brief Write a string to a worksheet cell.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param str     String to write to cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_string()` function writes a string to the cell
   * specified by `%row` and `%column`:
   *
   * @code
   *  worksheet.write_string(0, 0, "This phrase is English!");
   * @endcode
   *
   * @image html write_string01.png
   *
   * The `format` parameter is used to apply formatting to the cell. This
   * parameter can be `nullptr` to indicate no formatting or it can be a
   * @ref format.h "Format" object:
   *
   * @code
   *  xwpp::format* format = workbook.add_format();
   *  format->set_bold();
   *
   *  worksheet.write_string(0, 0, "This phrase is Bold!", format);
   * @endcode
   *
   * @image html write_string02.png
   *
   * Unicode strings are supported in UTF-8 encoding. This generally requires
   * that your source file is UTF-8 encoded or that the data has been read from
   * a UTF-8 source:
   *
   * @code
   *  worksheet.write_string(0, 0, "Это фраза на русском!");
   * @endcode
   *
   * @image html write_string03.png
   *
   * @todo Use overload of write (don't use suffix like "_string").
   */
  void write_string(row_num_t row_num, col_num_t col_num, const std::string& str, const format_t* format);
  /// @overload
  void write_string(row_num_t row_num, col_num_t col_num, const std::string& str);

  /**
   * @brief Write a number to a worksheet cell.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param number  The number to write to the cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_number()` function writes numeric types to the cell
   * specified by `%row` and `%column`:
   *
   * @code
   *  worksheet.write_number(0, 0, 123456);
   *  worksheet.write_number(1, 0, 2.3451);
   * @endcode
   *
   * @image html write_number01.png
   *
   * The native data type for all numbers in Excel is a IEEE-754 64-bit
   * double-precision floating point, which is also the default type used by
   * `%write_number()`.
   *
   * The `format` parameter is used to apply formatting to the cell. This
   * parameter can be `nullptr` to indicate no formatting or it can be a
   * @ref format.h "Format" object.
   *
   * @code
   *  xwpp::format_t* format = workbook.add_format();
   *  format->set_num_format("$#,##0.00");
   *
   *  worksheet.write_number(0, 0, 1234.567, format);
   * @endcode
   *
   * @image html write_number02.png
   *
   * @note Excel doesn't support `NaN`, `Inf` or `-Inf` as a number value. If
   * you are writing data that contains these values then your application
   * should convert them to a string or handle them in some other way.
   *
   * @todo Add overload for all integer and number types (template).
   */
  void write_number(row_num_t row_num, col_num_t col_num, double number, const format_t* format);
  /// @overload
  void write_number(row_num_t row_num, col_num_t col_num, double number);

  /**
   * @brief Write a date or time to a worksheet cell.
   *
   * @param row_num  The zero indexed row number.
   * @param col_num  The zero indexed column number.
   * @param datetime The datetime to write to the cell.
   * @param format   A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_datetime()` function can be used to write a date or
   * time to the cell specified by `%row` and `%column`:
   *
   * @dontinclude dates_and_times02.cpp
   * @skip include
   * @until num_format
   * @skip Feb
   * @until }
   *
   * The `format` parameter should be used to apply formatting to the cell using
   * a @ref format.h "Format" object as shown above. Without a date format the
   * datetime will appear as a number only.
   *
   * @see @ref working_with_dates for more information about handling dates and
   * times in Xlsxwriter++.
   *
   * @todo Add overload with other date and time type (including lxw_datetime).
   */
  void write_datetime(row_num_t row_num, col_num_t col_num, const std::chrono::system_clock::time_point& datetime,
                      const format_t* format);
  /// @overload
  void write_datetime(row_num_t row_num, col_num_t col_num, const std::chrono::system_clock::time_point& datetime);

  /**
   * @brief Write a Unix datetime to a worksheet cell.
   *
   * @param row_num  The zero indexed row number.
   * @param col_num  The zero indexed column number.
   * @param unixtime The Unix datetime to write to the cell.
   * @param format   A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_unixtime()` function can be used to write dates and
   * times in Unix date format to the cell specified by `%row` and
   * `%column`. [Unix Time](https://en.wikipedia.org/wiki/Unix_time) which is a
   * common integer time format. It is defined as the number of seconds since
   * the Unix epoch (1970-01-01 00:00 UTC). Negative values can also be used for
   * dates prior to 1970:
   *
   * @dontinclude dates_and_times03.cpp
   * @skip 1970
   * @until 2208988800
   *
   * The `format` parameter should be used to apply formatting to the cell using
   * a @ref format.h "Format" object as shown above. Without a date format the
   * datetime will appear as a number only.
   *
   * The output from this code sample is:
   *
   * @image html date_example03.png
   *
   * Unixtime is generally represented with a 32 bit `time_t` type which has a
   * range of approximately 1900-12-14 to 2038-01-19. To access the full Excel
   * date range of 1900-01-01 to 9999-12-31 this function uses a 64 bit
   * parameter.
   *
   * @see @ref working_with_dates for more information about handling dates and
   * times in Xlsxwriter++.
   */
  void write_unixtime(row_num_t row_num, col_num_t col_num, int64_t unixtime, const format_t* format);
  /// @overload
  void write_unixtime(row_num_t row_num, col_num_t col_num, int64_t unixtime);

  /**
   * @brief Write a formatted boolean worksheet cell.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param value   The boolean value to write to the cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * Write an Excel boolean to the cell specified by `%row` and `%column`:
   *
   * @code
   *  worksheet.write_boolean(2, 2, false, my_format);
   * @endcode
   */
  void write_boolean(row_num_t row_num, col_num_t col_num, bool value, const format_t* format);
  /// @overload
  void write_boolean(row_num_t row_num, col_num_t col_num, bool value);

  /**
   * @brief Write a formatted blank worksheet cell.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * Write a blank cell specified by `%row` and `%column`:
   *
   * @code
   *  worksheet.write_blank(1, 1, border_format);
   * @endcode
   *
   * This function is used to add formatting to a cell which doesn't contain a
   * string or number value.
   *
   * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
   * cell is a cell which doesn't contain data or formatting whilst a Blank cell
   * doesn't contain data but does contain formatting. Excel stores Blank cells
   * but ignores Empty cells.
   *
   * As such, if you write an empty cell without formatting it is ignored.
   */
  void write_blank(row_num_t row_num, col_num_t col_num, const format_t* format);

  /**
   * @brief Write a "Rich" multi-format string to a worksheet cell.
   *
   * @param row_num      The zero indexed row number.
   * @param col_num      The zero indexed column number.
   * @param rich_strings An array of format/string rich_string_tuple_t fragments.
   * @param format       A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_rich_string()` function is used to write strings with
   * multiple formats. For example to write the string 'This is **bold**
   * and this is *italic*' you would use the following:
   *
   * @code
   *  xwpp::format_t* bold = workbook.add_format();
   *  bold->set_bold();
   *
   *  xwpp::format_t* italic = workbook.add_format();
   *  italic->set_italic();
   *
   *  xwpp::rich_string_tuple_t fragment1 = {.str_ = "This is "};
   *  xwpp::rich_string_tuple_t fragment2 = {.format_ = bold, .str_ =  "bold"};
   *  xwpp::rich_string_tuple_t fragment3 = {.str_  = " and this is "};
   *  xwpp::rich_string_tuple_t fragment4 = {.format_ = italic, .str_ = "italic"};
   *
   *  std::vector<xwpp::rich_string_tuple_t> rich_string{
   *    fragment1, fragment2,
   *    fragment3, fragment4,
   *  };
   *
   *  worksheet.write_rich_string(CELL("A1"), rich_string, nullptr);
   * @endcode
   *
   * @image html rich_strings_small.png
   *
   * The basic rule is to break the string into fragments and put a `format_t`
   * object before the fragment that you want to format. So if we look at the
   * above example again:
   *
   * This is **bold** and this is *italic*
   *
   * The would be broken down into 4 fragments:
   *
   *      default: |This is |
   *      bold:    |bold|
   *      default: | and this is |
   *      italic:  |italic|
   *
   * This in then converted to the rich_string_tuple_t fragments shown in the
   * example above. For the default format we use `nullptr`.
   *
   * The fragments are passed to `%write_rich_string()` as a `std::vector<>`:
   *
   * @code
   *  std::vector<xwpp::rich_string_tuple> rich_string{
   *    fragment1, fragment2,
   *    fragment3, fragment4};
   *
   *  worksheet.write_rich_string(CELL("A1"), rich_string, nullptr);
   * @endcode
   *
   * @note Excel doesn't allow the use of two consecutive formats in a rich string or
   * an empty string fragment. For either of these conditions a warning is
   * raised and the input to `%write_rich_string()` is ignored.
   *
   * @todo Add API without `format`.
   */
  void write_rich_string(row_num_t row_num, col_num_t col_num, const std::vector<rich_string_tuple_t>& rich_strings,
                         const format_t* format);

  /**
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param url     The url to write to the cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_url()` function is used to write a URL/hyperlink to a
   * worksheet cell specified by `%row` and `%column`.
   *
   * @code
   *  worksheet.write_url(0, 0, "http://libxlsxwriter.github.io");
   * @endcode
   *
   * @image html hyperlinks_short.png
   *
   * The `format` parameter is used to apply formatting to the cell. This
   * parameter can be `nullptr`, in which case the default Excel blue underlined
   * hyperlink style will be used. If required a user defined @ref format.h
   * "Format" object can be used:
   *
   * @code
   *  xwpp::format_t* url_format = workbook.add_format();
   *
   *  url_format->set_underline(xwpp::format_underlines_t::SINGLE);
   *  url_format->set_font_color(xwpp::color_t::RED);
   * @endcode
   *
   * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
   * and `mailto:` :
   *
   * @code
   *  worksheet.write_url(0, 0, "ftp://www.python.org/");
   *  worksheet.write_url(1, 0, "http://www.python.org/");
   *  worksheet.write_url(2, 0, "https://www.python.org/");
   *  worksheet.write_url(3, 0, "mailto:jmcnamara@cpan.org");
   * @endcode
   *
   * An Excel hyperlink is comprised of two elements: the displayed string and
   * the non-displayed link. By default the displayed string is the same as the
   * link. However, it is possible to overwrite it with any other
   * `Xlsxwriter++` type using the appropriate `%write_*()`
   * function. The most common case is to overwrite the displayed link text with
   * another string. To do this we must also match the default URL format using
   * `workbook_t::get_default_url_format()`:
   *
   * @code
   *  // Write a hyperlink with the default blue underline format.
   *  worksheet.write_url(2, 0, "http://libxlsxwriter.github.io");
   *
   *  // Get the default url format.
   *  xwpp::format_t* url_format = workbook.get_default_url_format();
   *
   *  // Overwrite the hyperlink with a user defined string and default format.
   *  worksheet.write_string(2, 0, "Read the documentation.", url_format);
   * @endcode
   *
   * @image html hyperlinks_short2.png
   *
   * Two local URIs are supported: `internal:` and `external:`. These are used
   * for hyperlinks to internal worksheet references or external workbook and
   * worksheet references:
   *
   * @code
   *  worksheet.write_url(0, 0, "internal:Sheet2!A1");
   *  worksheet.write_url(1, 0, "internal:Sheet2!B2");
   *  worksheet.write_url(2, 0, "internal:Sheet2!A1:B2");
   *  worksheet.write_url(3, 0, "internal:'Sales Data'!A1");
   *  worksheet.write_url(4, 0, "external:c:\\temp\\foo.xlsx");
   *  worksheet.write_url(5, 0, "external:c:\\foo.xlsx#Sheet2!A1");
   *  worksheet.write_url(6, 0, "external:..\\foo.xlsx");
   *  worksheet.write_url( 7, 0, "external:..\\foo.xlsx#Sheet2!A1");
   *  worksheet.write_url(8, 0, "external:\\\\NET\\share\\foo.xlsx");
   * @endcode
   *
   * Worksheet references are typically of the form `Sheet1!A1`. You can also
   * link to a worksheet range using the standard Excel notation:
   * `Sheet1!A1:B2`.
   *
   * In external links the workbook and worksheet name must be separated by the
   * `#` character:
   *
   * @code
   *  worksheet.write_url(0, 0, "external:c:\\foo.xlsx#Sheet2!A1");
   * @endcode
   *
   * You can also link to a named range in the target worksheet: For example say
   * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
   * you could link to it as follows:
   *
   * @code
   *  worksheet.write_url(0, 0, "external:c:\\temp\\foo.xlsx#my_name");
   * @endcode
   *
   * Excel requires that worksheet names containing spaces or non alphanumeric
   * characters are single quoted as follows:
   *
   * @code
   *  worksheet.write_url(0, 0, "internal:'Sales Data'!A1");
   * @endcode
   *
   * Links to network files are also supported. Network files normally begin
   * with two back slashes as follows `\\NETWORK\etc`. In order to represent
   * this in a C string literal the backslashes should be escaped:
   *
   * @code
   *  worksheet.write_url(0, 0, "external:\\\\NET\\share\\foo.xlsx");
   * @endcode
   *
   * Alternatively, you can use Unix style forward slashes. These are
   * translated internally to backslashes:
   *
   * @code
   *  worksheet.write_url(0, 0, "external:c:/temp/foo.xlsx");
   *  worksheet.write_url(1, 0, "external://NET/share/foo.xlsx");
   * @endcode
   *
   * @note Xlsxwriter++ will escape the following characters in URLs as required
   * by Excel: `\s " < > \ [ ]  ^ { }`. Existing URL `%%xx` style escapes in
   * the string are ignored to allow for user-escaped strings.
   *
   * @note The maximum allowable URL length in recent versions of Excel is 2079
   * characters. In older versions of Excel (and Xlsxwriter++ <= 0.8.8) the
   * limit was 255 characters.
   *
   * @todo Document version with str and tooltip.
   */
  void write_url(row_num_t row_num, col_num_t col_num, const std::string& url, const format_t* format);
  /// @overload
  void write_url(row_num_t row_num, col_num_t col_num, const std::string& url);

  /* Don't document for now since the string option can be achieved by a
   * subsequent cell `write()` as shown in the docs, and the
   * tooltip option isn't very useful. */
  void write_url(row_num_t row_num, col_num_t col_num, const std::string& url, const format_t* format,
                 const std::string& str, const std::string& tooltip);

  /**
   * @brief Write a comment to a worksheet cell with options.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param text    The comment string to be written.
   * @param options Optional comment_options_t to control position and format
   *                of the comment.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_comment()` function is used to add a comment to a
   * cell. A comment is indicated in Excel by a small red triangle in the upper
   * right-hand corner of the cell. Moving the cursor over the red triangle will
   * reveal the comment.
   *
   * The following example shows how to add a comment to a cell:
   *
   * @code
   *  worksheet.write_comment(0, 0, "This is a comment");
   * @endcode
   *
   * @image html comments1.png
   *
   * The `%options` parameter can control the position, format and metadata of the
   * comment.
   *
   * The following example shows how to add a comment to a cell with options:
   *
   * @code
   *  xwpp::comment_options_t options = {.visible_ = LXW_COMMENT_DISPLAY_VISIBLE};
   *
   *  worksheet.write_comment(CELL("C6"), "Hello.", options);
   * @endcode
   *
   * The following options are available in `comment_options_t`:
   *
   * - `%author_`
   * - `%visible_`
   * - `%width_`
   * - `%height_`
   * - `%x_scale_`
   * - `%y_scale_`
   * - `%color_`
   * - `%font_name_`
   * - `%font_size_`
   * - `%start_row_`
   * - `%start_col_`
   * - `%x_offset_`
   * - `%y_offset_`
   *
   * @image html comments2.png
   *
   * Comment options are explained in detail in the @ref ww_comments_properties
   * section of the docs.
   */
  void write_comment(row_num_t row_num, col_num_t col_num, const std::string& text,
                     std::optional<comment_options_t> options);
  /// @overload
  void write_comment(row_num_t row_num, col_num_t col_num, const std::string& text);

  /**
   * @brief Write a formula to a worksheet cell.
   *
   * @param row     The zero indexed row number.
   * @param col     The zero indexed column number.
   * @param formula Formula string to write to cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_formula()` function writes a formula or function to
   * the cell specified by `%row` and `%column`:
   *
   * @code
   *  worksheet.write_formula(0, 0, "=B3 + 6");
   *  worksheet.write_formula(1, 0, "=SIN(PI()/4)");
   *  worksheet.write_formula(2, 0, "=SUM(A1:A2)");
   *  worksheet.write_formula(3, 0, "=IF(A3>1,\"Yes\", \"No\")");
   *  worksheet.write_formula(4, 0, "=AVERAGE(1, 2, 3, 4)");
   *  worksheet.write_formula(5, 0, "=DATEVALUE(\"1-Jan-2013\")");
   * @endcode
   *
   * @image html write_formula01.png
   *
   * The `%format` parameter is used to apply formatting to the cell. This
   * parameter can be `nullptr` to indicate no formatting or it can be a
   * @ref format.h "Format" object.
   *
   * `Xlsxwriter++` doesn't calculate the value of a formula and instead stores a
   * default value of `0`. The correct formula result is displayed in Excel, as
   * shown in the example above, since it recalculates the formulas when it loads
   * the file. For cases where this is an issue see the
   * `write_formula_num()` function and the discussion in that section.
   *
   * Formulas must be written with the US style separator/range operator which
   * is a comma (not semi-colon). Therefore a formula with multiple values
   * should be written as follows:
   *
   * @code
   *  // OK.
   *  worksheet.write_formula(0, 0, "=SUM(1, 2, 3)");
   *
   *  // NO. Error on load.
   *  worksheet.write_formula(1, 0, "=SUM(1; 2; 3)");
   * @endcode
   *
   * @see also @ref working_with_formulas.
   *
   * @todo Add overload with `%value` instead of other functions.
   */
  void write_formula(row_num_t row, col_num_t col, const std::string& formula, const format_t* format);
  /// @overload
  void write_formula(row_num_t row, col_num_t col, const std::string& formula);

  /**
   * @brief Write a formula to a worksheet cell with a user defined numeric
   * result.
   *
   * @param row     The zero indexed row number.
   * @param col     The zero indexed column number.
   * @param formula Formula string to write to cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   * @param result  A user defined numeric result for the formula.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_formula_num()` function writes a formula or Excel
   * function to the cell specified by `%row` and `%column` with a user defined
   * numeric result:
   *
   * @code
   *  // Required as a workaround only.
   *  worksheet.write_formula_num(0, 0, "=1 + 2", 3);
   * @endcode
   *
   * `Xlsxwriter++` doesn't calculate the value of a formula and instead stores
   * the value `0` as the formula result. It then sets a global flag in the XLSX
   * file to say that all formulas and functions should be recalculated when the
   * file is opened.
   *
   * This is the method recommended in the Excel documentation and in general it
   * works fine with spreadsheet applications.
   *
   * However, applications that don't have a facility to calculate formulas,
   * such as Excel Viewer, or some mobile applications will only display the `0`
   * results.
   *
   * If required, the `%write_formula_num()` function can be used to
   * specify a formula and its result.
   *
   * This function is rarely required and is only provided for compatibility
   * with some third party applications. For most applications the
   * `write_formula()` function is the recommended way of writing
   * formulas.
   *
   * @see @ref working_with_formulas.
   */
  void write_formula_num(row_num_t row, col_num_t col, const std::string& formula, const format_t* format,
                         double result);
  /// @overload
  void write_formula_num(row_num_t row, col_num_t col, const std::string& formula, double result);

  /**
   * @brief Write a formula to a worksheet cell with a user defined string
   * result.
   *
   * @param row     The zero indexed row number.
   * @param col     The zero indexed column number.
   * @param formula Formula string to write to cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   * @param result  A user defined string result for the formula.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_formula_str()` function writes a formula or Excel
   * function to the cell specified by `row` and `column` with a user defined
   * string result:
   *
   * @code
   *  // The example formula is A & B -> AB.
   *  worksheet.write_formula_str(0, 0, "=\"A\" & \"B\"", "AB");
   * @endcode
   *
   * The `%write_formula_str()` function is similar to the
   * `%write_formula_num()` function except it writes a string result
   * instead or a numeric result. See `write_formula_num()`  for more
   * details on why/when these functions are required.
   *
   * One place where the `%write_formula_str()` function may be required
   * is to specify an empty result which will force a recalculation of the formula
   * when loaded in LibreOffice.
   *
   * @code
   *  worksheet.write_formula_str(0, 0, "=Sheet1!$A$1", "");
   * @endcode
   *
   * @see the FAQ @ref faq_formula_zero.
   *
   * @see @ref working_with_formulas.
   */
  void write_formula_str(row_num_t row, col_num_t col, const std::string& formula, const format_t* format,
                         const std::string& result);
  /// @overload
  void write_formula_str(row_num_t row, col_num_t col, const std::string& formula, const std::string& result);

  /**
   * @brief Write an array formula to a worksheet cell.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   * @param formula   Array formula to write to cell.
   * @param format    A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_array_formula()` function writes an array formula to
   * a cell range. In Excel an array formula is a formula that performs a
   * calculation on a set of values.
   *
   * In Excel an array formula is indicated by a pair of braces around the
   * formula: `{=SUM(A1:B1*A2:B2)}`.
   *
   * Array formulas can return a single value or a range or values. For array
   * formulas that return a range of values you must specify the range that the
   * return values will be written to. This is why this function has `first_row`
   * `last_row`, `first_col` and `last_col` parameters.
   * The `RANGE()` macro can also be used to specify the range:
   *
   * @code
   *  worksheet.write_array_formula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}");
   *
   *  // Same as above using the RANGE() macro.
   *  worksheet.write_array_formula(RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}");
   * @endcode
   *
   * If the array formula returns a single value then the `first_` and `last_`
   * parameters should be the same:
   *
   * @code
   *  worksheet.write_array_formula(1, 0, 1, 0, "{=SUM(B1:C1*B2:C2)}");
   *  worksheet.write_array_formula(RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}");
   * @endcode
   *
   * @todo Add documentation of `%write_array_formula_num` (as `%write_array_formula_num` overload).
   */
  void write_array_formula(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                           const std::string& formula, const format_t* format);
  /// @overload
  void write_array_formula(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                           const std::string& formula);
  void write_array_formula_num(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                               const std::string& formula, const format_t* format, double result);

  /**
   * @brief Write an Excel 365 dynamic array formula to a worksheet range.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   * @param formula   Dynamic Array formula to write to cell.
   * @param format    A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_dynamic_array_formula()` function writes an Excel 365
   * dynamic array formula to a cell range. Some examples of functions that
   * return dynamic arrays are:
   *
   * - `FILTER`
   * - `RANDARRAY`
   * - `SEQUENCE`
   * - `SORTBY`
   * - `SORT`
   * - `UNIQUE`
   * - `XLOOKUP`
   * - `XMATCH`
   *
   * Dynamic array formulas and their usage in `Xlsxwriter++` is explained in
   * detail @ref ww_formulas_dynamic_arrays. The following is a example usage:
   *
   * @code
   *  worksheet.write_dynamic_array_formula(1, 5, 1, 5, "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)");
   * @endcode
   *
   * This formula gives the results shown in the image below.
   *
   * @image html dynamic_arrays02.png
   *
   * The need for the `_xlfn._xlws.` prefix in the formula is explained in @ref
   * ww_formulas_future.
   *
   * @todo Add documentation of `%write_dynamic_array_formula_num` (as overload of `%write_dynamic_array_formula`).
   */
  void write_dynamic_array_formula(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                                   const std::string& formula, const format_t* format);
  /// @overload
  void write_dynamic_array_formula(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                                   const std::string& formula);
  void write_dynamic_array_formula_num(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                                       const std::string& formula, const format_t* format, double result);

  /**
   * @brief Write an Excel 365 dynamic array formula to a worksheet cell.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param formula Formula string to write to cell.
   * @param format  A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%write_dynamic_formula()` function is similar to the
   * `write_dynamic_array_formula()` function, shown above, except
   * that it writes a dynamic array formula to a single cell, rather than a
   * range. This is a syntactic shortcut since the array range isn't generally
   * known for a dynamic range and specifying the initial cell is sufficient for
   * Excel, as shown in the example below:
   *
   * @code
   *  worksheet.write_dynamic_formula(7, 1, "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))");
   * @endcode
   *
   * This formula gives the following result:
   *
   * @image html dynamic_arrays01.png
   *
   * The need for the `_xlfn.` and `_xlfn._xlws.` prefixes in the formula is
   * explained in @ref ww_formulas_future.
   *
   * @todo Add documentation of `%write_dynamic_formula_num` (as overload of `%write_dynamic_formula`).
   */
  void write_dynamic_formula(row_num_t row_num, col_num_t col_num, const std::string& formula, const format_t* format);
  /// @overload
  void write_dynamic_formula(row_num_t row_num, col_num_t col_num, const std::string& formula);
  void write_dynamic_formula_num(row_num_t row, col_num_t col, const std::string& formula, const format_t* format,
                                 double result);

  /**
   * @brief Merge a range of cells.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   * @param str       String to write to the merged range.
   * @param format    A pointer to a Format instance or `nullptr`.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%merge_range()` function allows cells to be merged together
   * so that they act as a single area.
   *
   * Excel generally merges and centers cells at same time. To get similar
   * behavior with Xlsxwriter++ you need to apply a @ref format.h "Format"
   * object with the appropriate alignment:
   *
   * @code
   *  xwpp::format_t* merge_format = workbook.add_format();
   *  merge_format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
   *
   *  worksheet.merge_range(1, 1, 1, 3, "Merged Range", merge_format);
   * @endcode
   *
   * It is possible to apply other formatting to the merged cells as well:
   *
   * @code
   *  merge_format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
   *  merge_format->set_align(xwpp::format_alignments_t::VERTICAL_CENTER);
   *  merge_format->set_border(xwpp::format_borders_t::DOUBLE);
   *  merge_format->set_bold();
   *  merge_format->set_bg_color(xwpp::color_t(0xD7E4BC));
   *
   *  worksheet.merge_range(2, 1, 3, 3, "Merged Range", merge_format);
   * @endcode
   *
   * @image html merge.png
   *
   * The `%merge_range()` function writes a string using `write_string()`.
   * In order to write other data types, such as a number or a formula,
   * you can overwrite the first cell with a call to one of the other write
   * functions. The same Format should be used as was used in
   * the merged range.
   *
   * @code
   *  // First write a range with a blank string.
   *  worksheet.merge_range(1, 1, 1, 3, "", format);
   *
   *  // Then overwrite the first cell with a number.
   *  worksheet.write_number(1, 1, 123, format);
   * @endcode
   */
  void merge_range(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                   const std::string& str, const format_t* format);

  /**
   * @brief Add an Excel table to a worksheet.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   * @param options   A table_options_t struct to define the table options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%add_table()` function is used to add a table to a worksheet.
   * Tables in Excel are a way of grouping a range of cells into a
   * single entity that has common formatting or that can be referenced from
   * formulas. Tables can have column headers, autofilters, total rows, column
   * formulas and default formatting.
   *
   * @code
   *  worksheet.add_table(2, 1, 6, 5);
   * @endcode
   *
   * @image html tables1.png
   *
   * @see @ref working_with_tables for more detailed usage information
   */
  void add_table(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                 const std::optional<table_options_t>& options);
  /// @overload
  void add_table(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col);

  /**
   * @brief Insert an image in a worksheet cell, with options.
   *
   * @param row_num  The zero indexed row number.
   * @param col_num  The zero indexed column number.
   * @param filename The image filename, with path if required.
   * @param options  Optional image parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function can be used to insert a image into a worksheet. The image can
   * be in `PNG`, `JPEG`, `GIF` or `BMP` format:
   *
   * @code
   *  worksheet.insert_image(2, 1, "logo.png");
   * @endcode
   *
   * @image html insert_image.png
   *
   * **NOTE**:
   * The scaling of a image may be affected if is crosses a row that has its
   * default height changed due to a font that is larger than the default font
   * size or that has text wrapping turned on. To avoid this you should explicitly
   * set the height of the row using `set_row()` if it crosses an
   * inserted image. See @ref working_with_object_positioning.
   *
   * **NOTE on `SVG` files**:
   * Excel doesn't directly support SVG files in the same way as other image file
   * formats. It allows SVG to be inserted into a worksheet but converts them to,
   * and displays them as, PNG files. It stores the original SVG image in the file
   * so the original format can be retrieved. This removes the file size and
   * resolution advantage of using SVG files. As such SVG files are not supported
   * by `Xlsxwriter++` since a conversion to the PNG format would be required
   * and that format is already supported.
   *
   * BMP images are only supported for backward compatibility. In general it is
   * best to avoid `BMP` images since they aren't compressed. If used, `BMP` images
   * must be 24 bit, true color, bitmaps.
   *
   * `%insert_image()` function takes an optional
   * #image_options_t struct with the following members/options:
   *
   * - `%x_offset_`: Offset from the left of the cell in pixels.
   * - `%y_offset_`: Offset from the top of the cell in pixels.
   * - `%x_scale_`: X scale of the image as a decimal.
   * - `%y_scale_`: Y scale of the image as a decimal.
   * - `%object_position_`: See @ref working_with_object_positioning.
   * - `%description_`: Optional description or "Alt text" for the image.
   * - `%decorative_`: Optional parameter to mark image as decorative.
   * - `%url_`: Add an optional hyperlink to the image.
   * - `%tip_`: Add an optional mouseover tip for a hyperlink to the image.
   *
   * For example, to scale and position the image:
   *
   * @code
   *  xwpp::image_options_t options{
   *    .x_offset_ = 30,
   *    .y_offset_ = 10,
   *    .x_scale_  = 0.5,
   *    .y_scale_  = 0.5,
   *  };
   *
   *  worksheet.insert_image(2, 1, "logo.png", options);
   * @endcode
   *
   * @image html insert_image_opt.png
   *
   * The `%url_` field of `image_options_t` can be use to used to add a hyperlink
   * to an image:
   *
   * @code
   *  xwpp::image_options_t options{.url_ = "https://github.com/glerbret/xlsxwriterpp"};
   *
   *  worksheet.insert_image(3, 1, "logo.png", options);
   * @endcode
   *
   * The supported URL formats are the same as those supported by the
   * `write_url()` method and the same rules/limits apply.
   *
   * The `%tip_` field of `image_options_t` can be use to used to add a mouseover
   * tip to the hyperlink:
   *
   * @code
   *  xwpp::image_options_t options{.url_ = "https://github.com/glerbret/xlsxwriterpp",
   *                                .tip_ = "GitHub"};
   *  worksheet.insert_image(4, 1, "logo.png", options);
   * @endcode
   *
   * @note See the notes about row scaling and `BMP` images in
   * `insert_image()` above.
   *
   * @todo Add overload with path, ...
   */
  void insert_image(row_num_t row_num, col_num_t col_num, const std::string& filename,
                    std::optional<image_options_t> options);
  /// @overload
  void insert_image(row_num_t row_num, col_num_t col_num, const std::string& filename);

  /**
   * @brief Insert an image in a worksheet cell, from a memory buffer.
   *
   * @param row_num      The zero indexed row number.
   * @param col_num      The zero indexed column number.
   * @param image_buffer Vector of bytes that holds the image data.
   * @param options      Optional image parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function can be used to insert a image into a worksheet from a memory
   * buffer:
   *
   * @code
   *  worksheet.insert_image_buffer(CELL("B3"), image_buffer);
   * @endcode
   *
   * @image html image_buffer.png
   *
   * The buffer should be a pointer to an array of unsigned char data with a
   * specified size.
   *
   * `%insert_image_buffer()` function takes an optional
   * #image_options_t struct with the following members/options:
   *
   * - `%x_offset_`: Offset from the left of the cell in pixels.
   * - `%y_offset_`: Offset from the top of the cell in pixels.
   * - `%x_scale_`: X scale of the image as a decimal.
   * - `%y_scale_`: Y scale of the image as a decimal.
   * - `%object_position_`: See @ref working_with_object_positioning.
   * - `%description-`: Optional description or "Alt text" for the image.
   * - `%decorative-`: Optional parameter to mark image as decorative.
   * - `%url_`: Add an optional hyperlink to the image.
   * - `%tip_`: Add an optional mouseover tip for a hyperlink to the image.
   *
   * For example, to scale and position the image:
   *
   * @code
   *  xwpp::image_options_t options{.x_offset_ = 32, .y_offset_ = 4,
   *                                .x_scale_  = 2,  .y_scale_  = 1};
   *
   *  worksheet.insert_image_buffer(CELL("B3"), image_buffer, options);
   * @endcode
   *
   * @image html image_buffer_opt.png
   *
   * The buffer should be a pointer to a vector of unsigned char data.
   *
   * @see `insert_image_buffer()` for details about the supported
   * image formats, and other image options.
   */
  void insert_image_buffer(row_num_t row_num, col_num_t col_num, const std::vector<unsigned char>& image_buffer,
                           std::optional<image_options_t> options);
  /// @overload
  void insert_image_buffer(row_num_t row_num, col_num_t col_num, const std::vector<unsigned char>& image_buffer);

  /**
   * @brief Embed an image in a worksheet cell.
   *
   * @param row_num  The zero indexed row number.
   * @param col_num  The zero indexed column number.
   * @param filename The image filename, with path if required.
   * @param options  Optional image parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function can be used to embed a image into a worksheet cell and have the
   * image automatically scale to the width and height of the cell. The X/Y
   * scaling of the image is preserved but the size of the image is adjusted to
   * fit the largest possible width or height depending on the cell dimensions.
   *
   * This is the equivalent of Excel's menu option to insert an image using the
   * option to "Place in Cell" which is only available in Excel 365 versions from
   * 2023 onwards. For older versions of Excel a `#VALUE!` error is displayed.
   *
   * @dontinclude embed_images.cpp
   * @skip Change
   * @until B6
   *
   * @image html embed_image.png
   *
   * The `%embed_image()` function takes additional optional
   * image_options_t parameters, to add urls or format the cell background,
   * with the following members/options:
   *
   * - `%description_`: Optional description or "Alt text" for the image.
   * - `%decorative_`: Optional parameter to mark image as decorative.
   * - `%url_`: Add an optional hyperlink to the image.
   * - `%cell_format_`: Add a format for the cell behind the embedded image.
   *
   * @todo Doesn't work on LibreOffice, to investigate.
   */
  void embed_image(row_num_t row_num, col_num_t col_num, const std::string& filename,
                   std::optional<image_options_t> options);
  /// @overload
  void embed_image(row_num_t row_num, col_num_t col_num, const std::string& filename);

  /**
   * @brief Embed an image in a worksheet cell, from a memory buffer.
   *
   * @param row_num      The zero indexed row number.
   * @param col_num      The zero indexed column number.
   * @param image_buffer Vector of bytes that holds the image data.
   * @param options      Optional image parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function can be used to embed a image into a worksheet from a memory
   * buffer:
   *
   * @dontinclude embed_image_buffer.cpp
   * @skip Embed
   * @until B3
   *
   * @image html embed_image_buffer.png
   *
   * The `%embed_image_buffer()` function takes an optional
   * #image_options_t struct with the following members/options:
   *
   * - `%description_`: Optional description or "Alt text" for the image.
   * - `%decorative_`: Optional parameter to mark image as decorative.
   * - `%url_`: Add an optional hyperlink to the image.
   * - `%cell_format_`: Add a format for the cell behind the embedded image.
   *
   * See `embed_image()` for details about the supported image
   * formats, options, and other image features.
   */
  void embed_image_buffer(row_num_t row_num, col_num_t col_num, const std::vector<unsigned char>& image_buffer,
                          std::optional<image_options_t> options);
  /// @overload
  void embed_image_buffer(row_num_t row_num, col_num_t col_num, const std::vector<unsigned char>& image_buffer);

  /**
   * @brief Insert a chart object into a worksheet.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param chart   A xwpp::chart_t object created via xwpp::workbook_t::add_chart().
   * @param options Optional chart parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%insert_chart()` function can be used to insert a chart into
   * a worksheet. The chart object must be created first using the
   * `workbook_t::add_chart()` function and configured using the @ref chart.h
   * functions.
   *
   * @code
   *  // Create a chart object.
   *  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);
   *
   *  // Add a data series to the chart.
   *  chart.add_series("", "=Sheet1!$A$1:$A$6");
   *
   *  // Insert the chart into the worksheet.
   *  worksheet.insert_chart(0, 2, &chart);
   * @endcode
   *
   * @image html chart_working.png
   *
   * @note A chart may only be inserted into a worksheet once. If several
   * similar charts are required then each one must be created separately
   * with `%insert_chart()`.
   *
   * Optional `chart_options_t` parameter allow to scale and position the chart:
   *
   * @code
   *  xwpp::chart_options_t options{
   *    .x_offset_ = 30,
   *    .y_offset_ = 10,
   *    .x_scale_  = 0.5,
   *    .y_scale_  = 0.75,
   *  };
   *
   *  worksheet.insert_chart(0, 2, &chart, options);
   * @endcode
   *
   * @image html chart_line_opt.png
   */
  void insert_chart(row_num_t row_num, col_num_t col_num, chart_t* chart,
                    const std::optional<chart_options_t>& options);
  /// @overload
  void insert_chart(row_num_t row_num, col_num_t col_num, chart_t* chart); // TODO Ptr or ref

  /**
   * @brief Insert a button object into a worksheet.
   *
   * @param row_num The zero indexed row number.
   * @param col_num The zero indexed column number.
   * @param options A button_options_t object to set the button properties.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%insert_button()` function can be used to insert an Excel form
   * button into a worksheet. This function is generally only useful when
   * used in conjunction with the `workbook_t::add_vba_project()`
   * function to tie the button to a macro from an embedded VBA project:
   *
   * @code
   *  xwpp::button_options_t options{
   *    .caption_ = "Press Me",
   *    .macro_   = "say_hello",
   *  };
   *
   *  worksheet.insert_button(2, 1, options);
   * @endcode
   *
   * @image html macros.png
   *
   * The button properties are set using the `button_options_t` struct.
   *
   * @see @ref working_with_macros
   */
  void insert_button(row_num_t row_num, col_num_t col_num, const std::optional<button_options_t>& options);
  /// @overload
  void insert_button(row_num_t row_num, col_num_t col_num);

  /**
   * @brief Set the horizontal page breaks on a worksheet.
   *
   * @param breaks Array of page breaks.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_h_pagebreaks()` function adds horizontal page breaks to
   * a worksheet. A page break causes all the data that follows it to be printed
   * on the next page. Horizontal page breaks act between rows.
   *
   * The function takes a vector of one or more page breaks. The type of the
   * vector data is `row_num_t`:
   *
   * @code
   *  std::vector<xwpp::row_num_t> breaks1{20};
   *  std::vector<xwpp::row_num_t> breaks2{20, 40, 60, 80};
   *
   *  worksheet1.set_h_pagebreaks(breaks1);
   *  worksheet2.set_h_pagebreaks(breaks2);
   * @endcode
   *
   * To create a page break between rows 20 and 21 you must specify the break at
   * row 21. However in zero index notation this is actually row 20:
   *
   * @code
   *  // Break between row 20 and 21.
   *  std::vector<xwpp::row_num_t> breaks{20};
   *
   *  worksheet.set_h_pagebreaks(breaks);
   * @endcode
   *
   * @note There is an Excel limitation of 1023 horizontal page breaks per worksheet.
   *
   * @note If you specify the "fit to page" option via the
   * `fit_to_pages()` function it will override all manual page
   * breaks.
   */
  void set_h_pagebreaks(const std::vector<row_num_t>& breaks);

  /**
   * @brief Set the vertical page breaks on a worksheet.
   *
   * @param breaks Array of page breaks.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_v_pagebreaks()` function adds vertical page breaks to a
   * worksheet. A page break causes all the data that follows it to be printed
   * on the next page. Vertical page breaks act between columns.
   *
   * The function takes an vector of one or more page breaks. The type of the
   * vector data is `col_num_t`:
   *
   * @code
   *  std::vector<xwpp::col_num_t> breaks1{20};
   *  std::vector<xwpp::col_num_t> breaks2{20, 40, 60, 80};
   *
   *  worksheet1.set_v_pagebreaks(breaks1);
   *  worksheet2.set_v_pagebreaks(breaks2);
   * @endcode
   *
   * To create a page break between columns 20 and 21 you must specify the break
   * at column 21. However in zero index notation this is actually column 20:
   *
   * @code
   *  // Break between column 20 and 21.
   *  std::vector<xwpp::col_num_t> breaks{20};
   *
   *  worksheet.set_v_pagebreaks(breaks);
   * @endcode
   *
   * @note There is an Excel limitation of 1023 vertical page breaks per worksheet.
   *
   * @note If you specify the "fit to page" option via the
   * `fit_to_pages()` function it will override all manual page
   * breaks.
   */
  void set_v_pagebreaks(const std::vector<col_num_t>& breaks);

  /**
   * @brief Set the autofilter area in the worksheet.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%autofilter()` function allows an autofilter to be added to
   * a worksheet.
   *
   * An autofilter is a way of adding dropdown lists to the headers of a 2D
   * range of worksheet data. This allows users to filter the data based on
   * simple criteria so that some data is shown and some is hidden.
   *
   * @image html autofilter3.png
   *
   * To add an autofilter to a worksheet:
   *
   * @code
   *  worksheet.autofilter(0, 0, 50, 3);
   *
   *  // Same as above using the RANGE() macro.
   *  worksheet.autofilter(RANGE("A1:D51"));
   * @endcode
   *
   * In order to apply a filter condition it is necessary to add filter rules to
   * the columns using either the `%filter_column()`, `%filter_column2()` or
   * `%filter_list()` functions:
   *
   * - `filter_column()`: filter on a single criterion such as `"Column == East"`.
   * More complex conditions such as `<=` or `>=` can also be use.
   *
   * - `filter_column2()`: filter on two criteria such as `"Column == East or Column == West"`.
   * Complex conditions can also be used.
   *
   * - `filter_list()`: filter on a list of values such as `"Column in (East, West, North)"`.
   *
   * These functions are explained below. It isn't sufficient to just specify
   * the filter condition. You must also hide any rows that don't match the
   * filter condition.
   *
   * @see @ref ww_autofilters_data for more details.
   */
  void autofilter(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col);

  /**
   * @brief Write a filter rule to an autofilter column.
   *
   * @param col_num The column in the autofilter that the rule applies to.
   * @param rule    The autofilter rule.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%filter_column` function can be used to filter columns in a
   * autofilter range based on single rule conditions:
   *
   * @code
   *  xwpp::filter_rule_t filter_rule{
   *    .criteria_ = xwpp::filter_criteria_t::EQUAL_TO,
   *    .value_string_ = "East",
   *  };
   *
   *  worksheet.filter_column(0, filter_rule);
   * @endcode
   *
   * @image html autofilter4.png
   *
   * The rules and criteria are explained in more detail in @ref
   * ww_autofilters_criteria in @ref working_with_autofilters.
   *
   * The `col_num` parameter is a zero indexed column number and must refer to a
   * column in an existing autofilter created with `autofilter()`.
   *
   * It isn't sufficient to just specify the filter condition. You must also
   * hide any rows that don't match the filter condition.
   *
   * @see @ref ww_autofilters_data for more details.
   */
  void filter_column(col_num_t col_num, const filter_rule_t& rule);

  /**
   * @brief Write two filter rules to an autofilter column.
   *
   * @param col_num The column in the autofilter that the rules applies to.
   * @param rule1   First autofilter rule.
   * @param rule2   Second autofilter rule.
   * @param and_or  An and/or operator.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%filter_column2` function can be used to filter columns in a
   * autofilter range based on two rule conditions:
   *
   * @code
   *  xwpp::filter_rule_t filter_rule1{
   *    .criteria_ = xwpp::filter_criteria_t::EQUAL_TO,
   *    .value_string_ = "East",
   *  };
   *
   *  xwpp::filter_rule_t filter_rule2
   *  {
   *    .criteria_ = xwpp::filter_criteria_t::EQUAL_TO,
   *    .value_string_ = "South",
   *  };
   *
   *   worksheet.filter_column2(0, filter_rule1, filter_rule2, xwpp::filter_operator_t::OR);
   * @endcode
   *
   * @image html autofilter5.png
   *
   * The rules and criteria are explained in more detail in @ref
   * ww_autofilters_criteria in @ref working_with_autofilters.
   *
   * The `%col_num` parameter is a zero indexed column number and must refer to a
   * column in an existing autofilter created with `autofilter()`.
   *
   * The `%and_or` parameter is either "and" (`filter_operator_t::AND`) or "or"
   * (`filter_operator_t::OR`).
   *
   * It isn't sufficient to just specify the filter condition. You must also
   * hide any rows that don't match the filter condition.
   *
   * @see @ref ww_autofilters_data for more details.
   */
  void filter_column2(col_num_t col_num, const filter_rule_t& rule1, const filter_rule_t& rule2,
                      filter_operator_t and_or);

  /**
   * @brief Write multiple string filters to an autofilter column.
   *
   * @param col_num The column in the autofilter that the rules applies to.
   * @param list    A vector of strings to filter on.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%filter_column_list()` function can be used specify multiple
   * string matching criteria. This is a newer type of filter introduced in
   * Excel 2007. Prior to that it was only possible to have either 1 or 2 filter
   * conditions, such as the ones used by `filter_column()` and
   * `filter_column2()`.
   *
   * As an example, consider a column that contains data for the months of the
   * year. The `%filter_list()` function can be used to filter out
   * data rows for different months:
   *
   * @code
   *  std::vector<std::string> list{"March", "April", "May"};
   *
   *  worksheet.filter_list(0, list);
   * @endcode
   *
   * @image html autofilter2.png
   *
   * @code
   *  std::vector<std::string> list{"March", "April", "May", "Blanks"};
   *
   *  worksheet.filter_list(0, list);
   * @endcode
   *
   * It isn't sufficient to just specify the filter condition. You must also
   * hide any rows that don't match the filter condition.
   *
   * @see @ref ww_autofilters_data for more details.
   */
  void filter_list(col_num_t col_num, const std::vector<std::string>& list);

  /**
   * @brief Split and freeze a worksheet into panes.
   *
   * @param row_num The cell row (zero indexed).
   * @param col_num The cell column (zero indexed).
   *
   * The `%freeze_panes()` function can be used to divide a worksheet
   * into horizontal or vertical regions known as panes and to "freeze" these
   * panes so that the splitter bars are not visible.
   *
   * The parameters `%row_num` and `%col_num` are used to specify the location of the
   * split. It should be noted that the split is specified at the top or left of
   * a cell and that the function uses zero based indexing. Therefore to freeze
   * the first row of a worksheet it is necessary to specify the split at row 2
   * (which is 1 as the zero-based index).
   *
   * You can set one of the `%row_num` and `%col_num` parameters as zero if you do not
   * want either a vertical or horizontal split.
   *
   * @code
   *  worksheet1.freeze_panes(1, 0); // Freeze the first row.
   *  worksheet2.freeze_panes(0, 1); // Freeze the first column.
   *  worksheet3.freeze_panes(1, 1); // Freeze first row/column.
   * @endcode
   *
   * @todo Document the second API.
   */
  void freeze_panes(row_num_t row_num, col_num_t col_num);

  /* freeze_panes() with infrequent options. Undocumented for now. */
  void freeze_panes(row_num_t first_row, col_num_t first_col, row_num_t top_row, col_num_t left_col, bool type);

  /**
   * @brief Split a worksheet into panes.
   *
   * @param vertical   The position for the vertical split.
   * @param horizontal The position for the horizontal split.
   *
   * The `%split_panes()` function can be used to divide a worksheet
   * into horizontal or vertical regions known as panes. This function is
   * different from the `freeze_panes()` function in that the splits
   * between the panes will be visible to the user and each pane will have its
   * own scroll bars.
   *
   * The parameters `%vertical` and `%horizontal` are used to specify the vertical
   * and horizontal position of the split. The units for `%vertical` and
   * `%horizontal` are the same as those used by Excel to specify row height and
   * column width. However, the vertical and horizontal units are different from
   * each other. Therefore you must specify the `%vertical` and `%horizontal`
   * parameters in terms of the row heights and column widths that you have set
   * or the default values which are 15 for a row and 8.43 for a column.
   *
   * @code
   *  worksheet1.split_panes(15, 0);    // First row.
   *  worksheet2.split_panes(0,  8.43); // First column.
   *  worksheet3.split_panes(15, 8.43); // First row and column.
   * @endcode
   *
   * @todo Document the second API.
   */
  void split_panes(double vertical, double horizontal);

  /* split_panes() with infrequent options. Undocumented for now. */
  void split_panes(double vertical, double horizontal, row_num_t top_row, col_num_t left_col);

  /**
   * @brief Set the selected cell or cells in a worksheet:
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   *
   * The `%set_selection()` function can be used to specify which cell
   * or range of cells is selected in a worksheet: The most common requirement
   * is to select a single cell, in which case the `first_` and `last_`
   * parameters should be the same.
   *
   * The active cell within a selected range is determined by the order in which
   * `first_` and `last_` are specified.
   *
   * @code
   *  worksheet1.set_selection(3, 3, 3, 3);     // Cell D4.
   *  worksheet2.set_selection(3, 3, 6, 6);     // Cells D4 to G7.
   *  worksheet3.set_selection(6, 6, 3, 3);     // Cells G7 to D4.
   *  worksheet4.set_selection(RANGE("D4:G7")); // Using the RANGE macro.
   * @endcode
   */
  void set_selection(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col);

  /**
   * @brief Add a conditional format to a worksheet cell.
   *
   * @param row_num            The zero indexed row number.
   * @param col_num            The zero indexed column number.
   * @param conditional_format A `conditional_format_t` object to control the conditional format.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%conditional_format_cell()` function is used to set a
   * conditional format for a cell in a worksheet:
   *
   * @code
   *  conditional_format->type_     = xwpp::conditional_format_types_t::CELL;
   *  conditional_format->criteria_ = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO;
   *  conditional_format->value_    = 50;
   *  conditional_format->format_   = format;
   *
   *  worksheet.conditional_format_cell(CELL("A1"), conditional_format);
   * @endcode
   *
   * The conditional format parameters is specified in `conditional_format_t`.
   *
   * See @ref working_with_conditional_formatting for full details.
   */
  void conditional_format_cell(row_num_t row_num, col_num_t col_num, const conditional_format_t& conditional_format);

  /**
   * @brief Add a conditional format to a worksheet range.
   *
   * @param first_row          The first row of the range.
   * @param first_col          The first column of the range.
   * @param last_row           The last row of the range.
   * @param last_col           The last col of the range.
   * @param conditional_format A `conditional_format_t` object to control the conditional format.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%conditional_format_cell()` function is used to set a
   * conditional format for a range of cells in a worksheet:
   *
   * @code
   *  conditional_format->type_     = xwpp::conditional_format_types_t::CELL;
   *  conditional_format->criteria_ = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO;
   *  conditional_format->value_    = 50;
   *  conditional_format->format_   = format1;
   *
   *  worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format);
   *
   *  conditional_format->type_     = xwpp::conditional_format_types_t::CELL;
   *  conditional_format->criteria_ = xwpp::conditional_criteria_t::LESS_THAN;
   *  conditional_format->value_    = 50;
   *  conditional_format->format_   = format2;
   *
   *  worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format);
   * @endcode
   *
   * @image html conditional_format1.png
   *
   * The conditional format parameters is specified in `conditional_format_t`.
   *
   * @see @ref working_with_conditional_formatting for full details.
   */
  void conditional_format_range(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                                const conditional_format_t& conditional_format);

  /**
   * @brief Add a data validation to a cell.
   *
   * @param row_num    The zero indexed row number.
   * @param col_num    The zero indexed column number.
   * @param validation A data_validation_t object to control the validation.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%data_validation_cell()` function is used to construct an
   * Excel data validation or to limit the user input to a dropdown list of
   * values:
   *
   * @code
   *  xwpp::data_validation_t{
   *    .validate_       = xwpp::validation_types_t::INTEGER;
   *    .criteria_       = xwpp::validation_criteria_t::BETWEEN;
   *    .minimum_number_ = 1;
   *    .maximum_number_ = 10
   *  };
   *
   *  worksheet.data_validation_cell(2, 1, data_validation);
   *
   *  // Same as above with the CELL() macro.
   *  worksheet.data_validation_cell(CELL("B3"), data_validation);
   * @endcode
   *
   * @image html data_validate4.png
   *
   * Data validation and the various options of `data_validation_t` are
   * described in more detail in @ref working_with_data_validation.
   */
  void data_validation_cell(row_num_t row_num, col_num_t col_num, const data_validation_t& validation);

  /**
   * @brief Add a data validation to a range.
   *
   * @param first_row  The first row of the range.
   * @param first_col  The first column of the range.
   * @param last_row   The last row of the range.
   * @param last_col   The last col of the range.
   * @param validation A data_validation_t object to control the validation.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%data_validation_range()` function is the same as the
   * `data_validation_cell()`, see above, except the data validation
   * is applied to a range of cells:
   *
   * @code
   *  xwpp::data_validation_t{
   *    .validate_       = xwpp::validation_types_t::INTEGER;
   *    .criteria_       = xwpp::validation_criteria_t::BETWEEN;
   *    .minimum_number_ = 1;
   *    .maximum_number_ = 10
   *  };
   *
   *  worksheet.data_validation_range(2, 1, 4, 1, data_validation);
   *
   *  // Same as above with the RANGE() macro.
   *  worksheet.data_validation_range(RANGE("B3:B5"), data_validation);
   * @endcode
   *
   * Data validation and the various options of `data_validation_t` are
   * described in more detail in @ref working_with_data_validation.
   */
  void data_validation_range(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                             const data_validation_t& validation);

  /**
   * @brief Ignore various Excel errors/warnings in a worksheet for user
   * defined ranges.
   *
   * @param type  The type of error/warning to ignore. See `ignore_errors_t`.
   * @param range The range(s) for which the error/warning should be ignored.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%ignore_errors()` function can be used to ignore various
   * worksheet cell errors/warnings. For example the following code writes a
   * string that looks like a number:
   *
   * @code
   *  worksheet.write_string(CELL("D2"), "123");
   * @endcode
   *
   * This causes Excel to display a small green triangle in the top left hand
   * corner of the cell to indicate an error/warning:
   *
   * @image html ignore_errors1.png
   *
   * Sometimes these warnings are useful indicators that there is an issue in
   * the spreadsheet but sometimes it is preferable to turn them off. Warnings
   * can be turned off at the Excel level for all workbooks and worksheets by
   * using the using "Excel options -> Formulas -> Error checking
   * rules". Alternatively you can turn them off for individual cells in a
   * worksheet, or ranges of cells, using the `%ignore_errors()`
   * function with different ignore_errors_t options and ranges like this:
   *
   * @code
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "C3");
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::EVAL_ERROR, "C6");
   * @endcode
   *
   * The range can be a single cell, a range of cells, or multiple cells and
   * ranges separated by spaces:
   *
   * @code
   *  // Single cell.
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "C6");
   *
   *  // Or a single range:
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "C6:G8");
   *
   *  // Or multiple cells and ranges:
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "C6:E6 G1:G20 J2:J6");
   * @endcode
   *
   * @note Calling `%ignore_errors()` more than once for the same
   * `ignore_errors_t` type will overwrite the previous range.
   *
   * You can turn off warnings for an entire column by specifying the range from
   * the first cell in the column to the last cell in the column:
   *
   * @code
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "A1:A1048576");
   * @endcode
   *
   * Or for the entire worksheet by specifying the range from the first cell in
   * the worksheet to the last cell in the worksheet:
   *
   * @code
   *  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "A1:XFD1048576");
   * @endcode
   *
   * The worksheet errors/warnings that can be ignored are:
   *
   * - `ignore_errors_t::NUMBER_STORED_AS_TEXT`: Turn off errors/warnings for numbers
   *    stores as text.
   *
   * - `ignore_errors_t::EVAL_ERROR`: Turn off errors/warnings for formula errors (such
   *    as divide by zero).
   *
   * - `ignore_errors_t::FORMULA_DIFFERS`: Turn off errors/warnings for formulas that
   *    differ from surrounding formulas.
   *
   * - `ignore_errors_t::FORMULA_RANGE`: Turn off errors/warnings for formulas that
   *    omit cells in a range.
   *
   * - `ignore_errors_t::FORMULA_UNLOCKED`: Turn off errors/warnings for unlocked cells
   *    that contain formulas.
   *
   * - `ignore_errors_t::EMPTY_CELL_REFERENCE`: Turn off errors/warnings for formulas
   *    that refer to empty cells.
   *
   * - `ignore_errors_t::LIST_DATA_VALIDATION`: Turn off errors/warnings for cells in a
   *    table that do not comply with applicable data validation rules.
   *
   * - `ignore_errors_t::CALCULATED_COLUMN`: Turn off errors/warnings for cell formulas
   *    that differ from the column formula.
   *
   * - `ignore_errors_t::TWO_DIGIT_TEXT_YEAR`: Turn off errors/warnings for formulas
   *    that contain a two digit text representation of a year.
   */
  void ignore_errors(ignore_errors_t type, const std::string& range);

  /**
   * @brief Set the printed page header caption.
   *
   * @param str     The header string.
   * @param options Header options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * Headers and footers are generated using a string which is a combination of
   * plain text and control characters.
   *
   * The available control character are:
   *
   *   | Control         | Category      | Description           |
   *   | --------------- | ------------- | --------------------- |
   *   | `&L`            | Justification | Left                  |
   *   | `&C`            |               | Center                |
   *   | `&R`            |               | Right                 |
   *   | `&P`            | Information   | Page number           |
   *   | `&N`            |               | Total number of pages |
   *   | `&D`            |               | Date                  |
   *   | `&T`            |               | Time                  |
   *   | `&F`            |               | File name             |
   *   | `&A`            |               | Worksheet name        |
   *   | `&Z`            |               | Workbook path         |
   *   | `&fontsize`     | Font          | Font size             |
   *   | `&"font,style"` |               | Font name and style   |
   *   | `&U`            |               | Single underline      |
   *   | `&E`            |               | Double underline      |
   *   | `&S`            |               | Strikethrough         |
   *   | `&X`            |               | Superscript           |
   *   | `&Y`            |               | Subscript             |
   *   | `&G`            | Images        | Image placeholder     |
   *   | `&&`            | Miscellaneous | Literal ampersand &   |
   *
   * Text in headers and footers can be justified (aligned) to the left, center
   * and right by prefixing the text with the control characters `&L`, `&C` and
   * `&R`.
   *
   * For example (with ASCII art representation of the results):
   *
   * @code
   *  worksheet.set_header("&LHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    | Hello                                                         |
   *  //    |                                                               |
   *
   *
   *  worksheet.set_header("&CHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                          Hello                                |
   *  //    |                                                               |
   *
   *
   *  worksheet.set_header("&RHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                                                         Hello |
   *  //    |                                                               |
   * @endcode
   *
   * For simple text, if you do not specify any justification the text will be
   * centered. However, you must prefix the text with `&C` if you specify a font
   * name or any other formatting:
   *
   * @code
   *  worksheet.set_header("Hello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                          Hello                                |
   *  //    |                                                               |
   * @endcode
   *
   * You can have text in each of the justification regions:
   *
   * @code
   *  worksheet.set_header("&LCiao&CBello&RCielo");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    | Ciao                     Bello                          Cielo |
   *  //    |                                                               |
   * @endcode
   *
   * The information control characters act as variables that Excel will update
   * as the workbook or worksheet changes. Times and dates are in the users
   * default format:
   *
   * @code
   *  worksheet.set_header("&CPage &P of &N");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                        Page 1 of 6                            |
   *  //    |                                                               |
   *
   *  worksheet.set_header("&CUpdated at &T");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                    Updated at 12:30 PM                        |
   *  //    |                                                               |
   * @endcode
   *
   * You can specify the font size of a section of the text by prefixing it with
   * the control character `&n` where `n` is the font size:
   *
   * @code
   *  worksheet1.set_header("&C&30Hello Big");
   *  worksheet2.set_header("&C&10Hello Small");
   * @endcode
   *
   * You can specify the font of a section of the text by prefixing it with the
   * control sequence `&"font,style"` where `fontname` is a font name such as
   * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
   * "Courier New" or "Times New Roman" and `style` is one of the standard
   *
   * @code
   *  worksheet1.set_header("&C&\"Courier New,Italic\"Hello");
   *  worksheet2.set_header("&C&\"Courier New,Bold Italic\"Hello");
   *  worksheet3.set_header("&C&\"Times New Roman,Regular\"Hello");
   * @endcode
   *
   * It is possible to combine all of these features together to create
   * sophisticated headers and footers. As an aid to setting up complicated
   * headers and footers you can record a page set-up as a macro in Excel and
   * look at the format strings that VBA produces. Remember however that VBA
   * uses two double quotes `""` to indicate a single double quote. For the last
   * example above the equivalent VBA code looks like this:
   *
   * @code
   *  .LeftHeader = ""
   *  .CenterHeader = "&""Times New Roman,Regular""Hello"
   *  .RightHeader = ""
   * @endcode
   *
   * Alternatively you can inspect the header and footer strings in an Excel
   * file by unzipping it and grepping the XML sub-files. The following shows
   * how to do that using libxml's xmllint to format the XML for clarity:
   *
   * @code
   *  $ unzip myfile.xlsm -d myfile
   *  $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer" | sed 's/&amp;/\&/g'
   *
   *  <headerFooter scaleWithDoc="0">
   *    <oddHeader>&L&P</oddHeader>
   *  </headerFooter>
   * @endcode
   *
   * To include a single literal ampersand `&` in a header or footer you should
   * use a double ampersand `&&`:
   *
   * @code
   *  worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");
   * @endcode
   *
   * @note Excel requires that the header or footer string cannot be longer than 255
   * characters, including the control characters. Strings longer than this will
   * not be written.
   *
   * The `header_footer_options_t` options are:
   *
   * - `%margin_`: Header or footer margin in inches. The value must by larger
   *   than 0.0. The Excel default is 0.3.
   *
   * - `%image_left_`: The left header image filename, with path if required. This
   *   should have a corresponding `&G` placeholder in the `&L`
   *   section of the header/footer string.
   *
   * - `%image_center_`: The center header image filename, with path if
   *   required. This should have a corresponding `&G` placeholder in
   *   the `&C` section of the header/footer string.
   *
   * - `%image_right_`: The right header image filename, with path if
   *   required. This should have a corresponding `&G` placeholder in
   *   the `&R` section of the header/footer string.
   *
   * @code
   *  xwpp::header_footer_options_t header_options = { .margin_ = 0.2 };
   *
   *  worksheet.set_header("Some text", header_options);
   * @endcode
   *
   * Images can be inserted in the header by specifying the `&G`
   * placeholder and a filename/path to the image:
   *
   * @code
   *  xwpp::header_footer_options_t header_options = {.image_left_ = "logo.png"};
   *
   *  worksheet1set_header("&L&G", header_options);
   * @endcode
   *
   * @image html headers_footers.png
   */
  void set_header(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_header(const std::string& str);

  /**
   * @brief Set the printed page footer caption with additional options.
   *
   * @param str     The footer string.
   * @param options Footer options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The syntax of this function is the same as `set_header()`.
   */
  void set_footer(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_footer(const std::string& str);

  /**
   * @brief Set the worksheet margins for the printed page.
   *
   * @param left   Left margin in inches.   Excel default is 0.7.
   * @param right  Right margin in inches.  Excel default is 0.7.
   * @param top    Top margin in inches.    Excel default is 0.75.
   * @param bottom Bottom margin in inches. Excel default is 0.75.
   *
   * The `%set_margins()` function is used to set the margins of the
   * worksheet when it is printed. The units are in inches. Specifying `-1` for
   * any parameter will give the default Excel value as shown above.
   *
   * @code
   *  worksheet.set_margins(1.3, 1.2, -1, -1);
   * @endcode
   */
  void set_margins(double left, double right, double top, double bottom);

  /**
   * @brief Set the color of the worksheet tab.
   *
   * @param color The tab color.
   *
   * The `%set_tab_color()` function is used to change the color of
   * the worksheet tab:
   *
   * @code
   *  worksheet1.set_tab_color(xwpp::color_t::RED);
   *  worksheet2.set_tab_color(xwpp::color_t::GREEN);
   *  worksheet3.set_tab_color(xwpp::color_t(0xFF9900)); // Orange.
   * @endcode
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   */
  void set_tab_color(color_t color);

  /**
   * @brief Set the worksheet zoom factor.
   *
   * @param scale Worksheet zoom factor.
   *
   * Set the worksheet zoom factor in the range `10 <= zoom <= 400`:
   *
   * @code
   *  worksheet1.set_zoom(50);
   *  worksheet2.set_zoom(75);
   *  worksheet3.set_zoom(300);
   *  worksheet4.set_zoom(400);
   * @endcode
   *
   * The default zoom factor is 100. It isn't possible to set the zoom to
   * "Selection" because it is calculated by Excel at run-time.
   *
   * @note `%set_zoom()` does not affect the scale of the printed
   * page. For that you should use `set_print_scale()`.
   */
  void set_zoom(uint16_t scale);

  /**
   * @brief Set the default row properties.
   *
   * @param height           Default row height.
   * @param hide_unused_rows Hide unused cells.
   *
   * The `%set_default_row()` function is used to set Excel default
   * row properties such as the default height and the option to hide unused
   * rows. These parameters are an optimization used by Excel to set row
   * properties without generating a very large file with an entry for each row.
   *
   * To set the default row height:
   *
   * @code
   *  worksheet.set_default_row(24, false);
   * @endcode
   *
   * To hide unused rows:
   *
   * @code
   *  worksheet.set_default_row(15, true);
   * @endcode
   *
   * @note In the previous case we use the default height `DEF_ROW_HEIGHT`
   * so the the height remains unchanged.
   */
  void set_default_row(double height, bool hide_unused_rows);

  /**
   * @brief Set the background image for a worksheet.
   *
   * @param filename The image filename, with path if required.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_background()` function can be used to set the
   * background image for a worksheet:
   *
   * @code
   *  worksheet.set_background("logo.png");
   * @endcode
   *
   * @image html background.png
   *
   * The `%set_background()` method supports all the image formats supported by
   * `insert_image()`.
   *
   * Some people use this method to add a watermark background to their
   * document. However, Microsoft recommends using a header image
   * [to set a watermark][watermark]. The choice of method depends on whether
   * you want the watermark to be visible in normal viewing mode or just
   * when the file is printed. In Xlsxwriter++ you can get the header watermark
   * effect using `set_header()`:
   *
   * @code
   *  xwpp::header_footer_options_t header_options{.image_center_ = "watermark.png"};
   *  worksheet.set_header(worksheet, "&C&G", header_options);
   * @endcode
   *
   * [watermark]:https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
   *
   * @todo Add API with std::filesystem::path.
   */
  void set_background(const std::string& filename);

  /**
   * @brief Set the background image for a worksheet, from a buffer.
   *
   * @param image_buffer Vector of bytes that holds the image data.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function can be used to insert a background image into a worksheet
   * from a memory buffer:
   *
   * @code
   *  worksheet.set_background_buffer(image_buffer);
   * @endcode
   *
   * The buffer should be a pointer to an array of unsigned char data with a
   * specified size.
   *
   * @see `set_background()` for more details.
   */
  void set_background_buffer(const std::vector<unsigned char>& image_buffer);

  /**
   * @brief Set the VBA name for the worksheet.
   *
   * @param name Name of the worksheet used by VBA.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_vba_name()` function can be used to set the VBA name for
   * the worksheet. This is sometimes required when a vbaProject macro included
   * via `workbook_t::add_vba_project()` refers to the worksheet by a name other
   * than the worksheet name:
   *
   * @code
   *  workbook.set_vba_name("MyWorkbook");
   *  worksheet.set_vba_name("MySheet1");
   * @endcode
   *
   * In general Excel uses the worksheet name such as `"Sheet1"` as the VBA name.
   * However, this can be changed in the VBA environment or if the the macro was
   * extracted from a foreign language version of Excel.
   *
   * @see @ref working_with_macros
   */
  void set_vba_name(const std::string& name);

  /**
   * @brief Set the default author of the cell comments.
   *
   * @param author The name of the comment author.
   *
   * This `%set_comments_author()` function is used to set the
   * default author of all cell comments:
   *
   * @code
   *  worksheet.set_comments_author("John Doe")
   * @endcode
   *
   * Individual authors can be set using the `%author_` option of the
   * `comment_options_t` struct and the `write_comment()`
   * function (see above and @ref ww_comments_author).
   */
  void set_comments_author(const std::string& author);

  /**
   * @brief Set current worksheet as the first visible sheet tab.
   *
   * The `activate()` function determines which worksheet is initially
   * selected.  However, if there are a large number of worksheets the selected
   * worksheet may not appear on the screen. To avoid this you can select the
   * leftmost visible worksheet tab using `%set_first_sheet()`:
   *
   * @code
   *  worksheet19.set_first_sheet(); // First visible worksheet tab.
   *  worksheet20.activate();        // First visible worksheet.
   * @endcode
   *
   * This function is not required very often. The default value is the first
   * worksheet.
   */
  void set_first_sheet();

  /**
   * @brief Set the first visible cell at the top left of a worksheet.
   *
   * @param row_num The cell row (zero indexed).
   * @param col_num The cell column (zero indexed).
   *
   * The `%set_top_left_cell()` function can be used to set the
   * top leftmost visible cell in the worksheet:
   *
   * @code
   *  worksheet.set_top_left_cell(31, 26);
   *  worksheet.set_top_left_cell(CELL("AA32")); // Same as above.
   * @endcode
   *
   * @image html top_left_cell.png
   */
  void set_top_left_cell(row_num_t row_num, col_num_t col_num);

  /**
   * @brief Set the paper type for printing.
   *
   * @param paper_type The Excel paper format type.
   *
   * This function is used to set the paper format for the printed output of a
   * worksheet. The following paper styles are available:
   *
   *   Index    | Paper format            | Paper size
   *   :------- | :---------------------- | :-------------------
   *   0        | Printer default         | Printer default
   *   1        | Letter                  | 8 1/2 x 11 in
   *   2        | Letter Small            | 8 1/2 x 11 in
   *   3        | Tabloid                 | 11 x 17 in
   *   4        | Ledger                  | 17 x 11 in
   *   5        | Legal                   | 8 1/2 x 14 in
   *   6        | Statement               | 5 1/2 x 8 1/2 in
   *   7        | Executive               | 7 1/4 x 10 1/2 in
   *   8        | A3                      | 297 x 420 mm
   *   9        | A4                      | 210 x 297 mm
   *   10       | A4 Small                | 210 x 297 mm
   *   11       | A5                      | 148 x 210 mm
   *   12       | B4                      | 250 x 354 mm
   *   13       | B5                      | 182 x 257 mm
   *   14       | Folio                   | 8 1/2 x 13 in
   *   15       | Quarto                  | 215 x 275 mm
   *   16       | ---                     | 10x14 in
   *   17       | ---                     | 11x17 in
   *   18       | Note                    | 8 1/2 x 11 in
   *   19       | Envelope 9              | 3 7/8 x 8 7/8
   *   20       | Envelope 10             | 4 1/8 x 9 1/2
   *   21       | Envelope 11             | 4 1/2 x 10 3/8
   *   22       | Envelope 12             | 4 3/4 x 11
   *   23       | Envelope 14             | 5 x 11 1/2
   *   24       | C size sheet            | ---
   *   25       | D size sheet            | ---
   *   26       | E size sheet            | ---
   *   27       | Envelope DL             | 110 x 220 mm
   *   28       | Envelope C3             | 324 x 458 mm
   *   29       | Envelope C4             | 229 x 324 mm
   *   30       | Envelope C5             | 162 x 229 mm
   *   31       | Envelope C6             | 114 x 162 mm
   *   32       | Envelope C65            | 114 x 229 mm
   *   33       | Envelope B4             | 250 x 353 mm
   *   34       | Envelope B5             | 176 x 250 mm
   *   35       | Envelope B6             | 176 x 125 mm
   *   36       | Envelope                | 110 x 230 mm
   *   37       | Monarch                 | 3.875 x 7.5 in
   *   38       | Envelope                | 3 5/8 x 6 1/2 in
   *   39       | Fanfold                 | 14 7/8 x 11 in
   *   40       | German Std Fanfold      | 8 1/2 x 12 in
   *   41       | German Legal Fanfold    | 8 1/2 x 13 in
   *
   * @code
   *  worksheet1.set_paper(1);  // US Letter
   *  worksheet2.set_paper(9);  // A4
   * @endcode
   *
   * If you do not specify a paper type the worksheet will print using the
   * printer's default paper style.
   *
   * @note It is likely that not all of these paper types will be available to
   * the end user since it will depend on the paper formats that the user's
   * printer supports. Therefore, it is best to stick to standard paper types:
   *
   * @todo Use enum as `paper_type`.
   */
  void set_paper(uint8_t paper_type);

  /**
   * @brief Set the page orientation as landscape.
   *
   * This function is used to set the orientation of a worksheet's printed page
   * to landscape:
   *
   * @code
   *  worksheet.set_landscape();
   * @endcode
   */
  void set_landscape();

  /**
   * @brief Set the page orientation as portrait.
   *
   * This function is used to set the orientation of a worksheet's printed page
   * to portrait. The default worksheet orientation is portrait, so this
   * function isn't generally required:
   *
   * @code
   *  worksheet.set_portrait();
   * @endcode
   */
  void set_portrait();

  /**
   * @brief Set the page layout to page view mode.
   *
   * This function is used to display the worksheet in "Page View/Layout" mode:
   *
   * @code
   *  worksheet.set_page_view();
   * @endcode
   */
  void set_page_view();

  // TODO Add documentation (for test only)
  void set_dpi(uint16_t horizontal_dpi, uint16_t vertical_dpi);

  /**
   * @brief Make all comments in the worksheet visible.
   *
   * This `%show_comments()` function is used to make all cell
   * comments visible when a worksheet is opened:
   *
   * @code
   *  worksheet.show_comments();
   * @endcode
   *
   * Individual comments can be made visible or hidden using the `%visible_`
   * option of the `comment_options_t` struct and the `write_comment()` function
   * (see above and @ref  ww_comments_visible).
   */
  void show_comments();

  /**
   * @brief Display the worksheet cells from right to left for some versions of
   * Excel.
   *
   * The `%right_to_left()` function is used to change the default
   * direction of the worksheet from left-to-right, with the `A1` cell in the
   * top left, to right-to-left, with the `A1` cell in the top right.
   *
   * @code
   *  worksheet.right_to_left();
   * @endcode
   *
   * This is useful when creating Arabic, Hebrew or other near or far eastern
   * worksheets that use right-to-left as the default direction.
   */
  void right_to_left();

  /**
   * @brief Hide zero values in worksheet cells.
   *
   * The `%hide_zero()` function is used to hide any zero values that
   * appear in cells:
   *
   * @code
   *  worksheet.hide_zero();
   * @endcode
   */
  void hide_zero();

  /**
   * @brief Make a worksheet the active, i.e., visible worksheet.
   *
   * The `%activate()` function is used to specify which worksheet is
   * initially visible in a multi-sheet workbook:
   *
   * @code
   *  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
   *  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
   *  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
   *
   *  worksheet3.activate();
   * @endcode
   *
   * @image html worksheet_activate.png
   *
   * More than one worksheet can be selected via the `select()`
   * function, see below, however only one worksheet can be active.
   *
   * The default active worksheet is the first worksheet.
   */
  void activate();

  /**
   * @brief Set a worksheet tab as selected.
   *
   * The `%select()` function is used to indicate that a worksheet is
   * selected in a multi-sheet workbook:
   *
   * @code
   *  worksheet1.activate();
   *  worksheet2.select();
   *  worksheet3.select();
   * @endcode
   *
   * A selected worksheet has its tab highlighted. Selecting worksheets is a
   * way of grouping them together so that, for example, several worksheets
   * could be printed in one go. A worksheet that has been activated via the
   * `activate()` function will also appear as selected.
   */
  void select();

  /**
   * @brief Protect elements of a worksheet from modification.
   *
   * @param password A worksheet password.
   * @param options  Worksheet elements to protect.
   *
   * The `%protect()` function protects worksheet elements from
   * modification:
   *
   * @code
   *  worksheet.protect("Some Password", options);
   * @endcode
   *
   * The `password` and `options` are both optional:
   *
   * @code
   *  worksheet1.protect();
   *  worksheet2.protect(my_options);
   *  worksheet3.protect("password");
   *  worksheet4.protect("password", my_options);
   * @endcode
   *
   * Passing no password (or an empty one) is the same as turning on protection
   * without a password. Passing no password and no options, or any other
   * combination has the effect of enabling a cell's `locked` and `hidden`
   * properties if they have been set.
   *
   * A *locked* cell cannot be edited and this property is on by default for all
   * cells. A *hidden* cell will display the results of a formula but not the
   * formula itself. These properties can be set using the `format_t::set_unlocked()`
   * and `format_t::set_hidden()` format functions.
   *
   * You can specify which worksheet elements you wish to protect by passing a
   * `protection_t` pointer in the `options` argument with any or all of the
   * following members set:
   *
   * - `%no_select_locked_cells_`
   * - `%no_select_unlocked_cells_`
   * - `%format_cells_`
   * - `%format_columns_`
   * - `%format_rows_`
   * - `%insert_columns_`
   * - `%insert_rows_`
   * - `%insert_hyperlinks_`
   * - `%delete_columns_`
   * - `%delete_rows_`
   * - `%sort_`
   * - `%autofilter_`
   * - `%pivot_tables_`
   * - `%scenarios_`
   * - `%objects_`
   *
   * All parameters are off by default. Individual elements can be protected as
   * follows:
   *
   * @code
   *  xwpp::protection_t options{
   *    .format_cells_      = true,
   *    .insert_hyperlinks_ = true,
   *    .insert_rows_       = true,
   *    .delete_rows_       = true,
   *    .insert_columns_    = true,
   *    .delete_columns_    = true,
   *  };
   *
   *  worksheet.protect(worksheet, options);
   * @endcode
   *
   * @see the `format_t::set_unlocked()` and `format_t::set_hidden()` format functions.
   *
   * @note Sheet level passwords in Excel offer **very** weak
   * protection. They don't encrypt your data and are very easy to
   * deactivate. Full workbook encryption is not supported by `Xlsxwriter++`
   * since it requires a completely different file format.
   */
  void protect(const std::string& password, std::optional<protection_t> options);
  /// @overload
  void protect(const std::string& password);
  /// @overload
  void protect(std::optional<protection_t> options);
  /// @overload
  void protect();

  /**
   * @brief Hide the current worksheet.
   *
   * The `%hide()` function is used to hide a worksheet:
   *
   * @code
   *  worksheet.hide();
   * @endcode
   *
   * You may wish to hide a worksheet in order to avoid confusing a user with
   * intermediate data or calculations.
   *
   * @image html hide_sheet.png
   *
   * A hidden worksheet can not be activated or selected so this function is
   * mutually exclusive with the `activate()` and `select()` functions.
   * In addition, since the first worksheet will default to being the
   * active worksheet, you cannot hide the first worksheet without activating
   * another sheet:
   *
   * @code
   *  worksheet2.activate(worksheet2);
   *  worksheet1.hide(worksheet1);
   * @endcode
   */
  void hide();

  /**
   * @brief Set the option to display or hide gridlines on the screen and
   * the printed page.
   *
   * @param option Gridline option.
   *
   * Display or hide screen and print gridlines using one of the values of
   * `gridlines_t`.
   *
   * @code
   *  worksheet1.gridlines(worksheet1, xwpp::gridlines_t::HIDE_ALL_GRIDLINES);
   *
   *  worksheet2.gridlines(worksheet2, xwpp::gridlines_t::SHOW_PRINT_GRIDLINES);
   * @endcode
   *
   * The Excel default is that the screen gridlines are on  and the printed
   * worksheet is off.
   */
  void gridlines(gridlines_t option);

  /**
   * @brief Set the Outline and Grouping display properties.
   *
   * @param visible       Outlines are visible. Optional, defaults to True.
   * @param symbols_below Show row outline symbols below the outline bar.
   * @param symbols_right Show column outline symbols to the right of outline.
   * @param auto_style    Use Automatic outline style.
   *
   * The `%outline_settings()` method is used to control the
   * appearance of outlines in Excel. Outlines are described the section on
   * @ref working_with_outlines.
   *
   * The `%visible` parameter is used to control whether or not outlines are
   * visible. Setting this parameter to False will cause all outlines on the
   * worksheet to be hidden. They can be un-hidden in Excel by means of the
   * "Show Outline Symbols" command button. The default Excel setting is True
   * for visible outlines.
   *
   * The `%symbols_below` parameter is used to control whether the row outline
   * symbol will appear above or below the outline level bar. The default Excel
   * setting is True for symbols to appear below the outline level bar.
   *
   * The `%symbols_right` parameter is used to control whether the column outline
   * symbol will appear to the left or the right of the outline level bar. The
   * default Excel setting is True for symbols to appear to the right of the
   * outline level bar.
   *
   * The `%auto_style` parameter is used to control whether the automatic outline
   * generator in Excel uses automatic styles when creating an outline. This has
   * no effect on a file generated by XlsxWriter but it does have an effect on
   * how the worksheet behaves after it is created. The default Excel setting is
   * False for "Automatic Styles" to be turned off.
   *
   * The default settings for all of these parameters in `Xlsxwriter++`
   * correspond to Excel's default parameters and are shown below:
   *
   * @code
   *  worksheet.outline_settings(true, true, true, false);
   * @endcode
   *
   * The worksheet parameters controlled by `%outline_settings()` are
   * rarely used.
   */
  void outline_settings(bool visible, bool symbols_below, bool symbols_right, bool auto_style);

  /**
   * @brief Set the print area for a worksheet.
   *
   * @param first_row The first row of the range.
   * @param first_col The first column of the range.
   * @param last_row  The last row of the range.
   * @param last_col  The last col of the range.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function is used to specify the area of the worksheet that will be
   * printed. The `RANGE()` macro is often convenient for this.
   *
   * @code
   *  worksheet.print_area(0, 0, 41, 10); // A1:K42.
   *
   *  // Same as:
   *  worksheet.print_area(RANGE("A1:K42"));
   * @endcode
   *
   * In order to set a row or column range you must specify the entire range:
   *
   * @code
   *  worksheet.print_area(RANGE("A1:H1048576")); // Same as A:H.
   * @endcode
   */
  void print_area(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col);

  /**
   * @brief Fit the printed area to a specific number of pages both vertically
   * and horizontally.
   *
   * @param width  Number of pages horizontally.
   * @param height Number of pages vertically.
   *
   * The `%fit_to_pages()` function is used to fit the printed area to
   * a specific number of pages both vertically and horizontally. If the printed
   * area exceeds the specified number of pages it will be scaled down to
   * fit. This ensures that the printed area will always appear on the specified
   * number of pages even if the page size or margins change:
   *
   * @code
   *  worksheet1.fit_to_pages(1, 1); // Fit to 1x1 pages.
   *  worksheet2.fit_to_pages(2, 1); // Fit to 2x1 pages.
   *  worksheet3.fit_to_pages(1, 2); // Fit to 1x2 pages.
   * @endcode
   *
   * The print area can be defined using the `print_area()` function
   * as described above.
   *
   * A common requirement is to fit the printed output to `n` pages wide but
   * have the height be as long as necessary. To achieve this set the `height`
   * to zero:
   *
   * @code
   *  // 1 page wide and as long as necessary.
   *  worksheet.fit_to_pages(1, 0);
   * @endcode
   *
   * @note Although it is valid to use both `%fit_to_pages()` and
   * `%set_print_scale()` on the same worksheet Excel only allows one
   * of these options to be active at a time. The last function call made will
   * set the active option.
   *
   * @note The `%fit_to_pages()` function will override any manual page
   * breaks that are defined in the worksheet.
   *
   * @note When using `%fit_to_pages()` it may also be required to set the
   * printer paper size using `%set_paper()` or else Excel will
   * default to "US Letter".
   */
  void fit_to_pages(uint16_t width, uint16_t height);

  /**
   * @brief Set the order in which pages are printed.
   *
   * The `%print_across()` function is used to change the default
   * print direction. This is referred to by Excel as the sheet "page order":
   *
   * @code
   *  worksheet.print_across();
   * @endcode
   *
   * The default page order is shown below for a worksheet that extends over 4
   * pages. The order is called "down then across":
   *
   *     [1] [3]
   *     [2] [4]
   *
   * However, by using the `print_across` function the print order will be
   * changed to "across then down":
   *
   *     [1] [2]
   *     [3] [4]
   */
  void print_across();

  /**
   * @brief Center the printed page horizontally.
   *
   * Center the worksheet data horizontally between the margins on the printed
   * page:
   *
   * @code
   *  worksheet.center_horizontally();
   * @endcode
   */
  void center_horizontally();

  /**
   * @brief Center the printed page vertically.
   *
   * Center the worksheet data vertically between the margins on the printed
   * page:
   *
   * @code
   *  worksheet.center_vertically();
   * @endcode
   */
  void center_vertically();

  /**
   * @brief Set the option to print the row and column headers on the printed
   * page.
   *
   * When printing a worksheet from Excel the row and column headers (the row
   * numbers on the left and the column letters at the top) aren't printed by
   * default.
   *
   * This function sets the printer option to print these headers:
   *
   * @code
   *  worksheet.print_row_col_headers();
   * @endcode
   */
  void print_row_col_headers();

  /**
   * @brief Set the number of rows to repeat at the top of each printed page.
   *
   * @param first_row First row of repeat range.
   * @param last_row  Last row of repeat range.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * For large Excel documents it is often desirable to have the first row or
   * rows of the worksheet print out at the top of each page.
   *
   * This can be achieved by using this function. The parameters `%first_row`
   * and `%last_row` are zero based:
   *
   * @code
   *  worksheet1.repeat_rows(0, 0); // Repeat the first row.
   *  worksheet2.repeat_rows(0, 1); // Repeat the first two rows.
   * @endcode
   */
  void repeat_rows(row_num_t first_row, row_num_t last_row);

  /**
   * @brief Set the number of columns to repeat at the top of each printed page.
   *
   * @param first_col First column of repeat range.
   * @param last_col  Last column of repeat range.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * For large Excel documents it is often desirable to have the first column or
   * columns of the worksheet print out at the left of each page.
   *
   * This can be achieved by using this function. The parameters `%first_col`
   * and `%last_col` are zero based:
   *
   * @code
   *  worksheet1.repeat_columns(0, 0); // Repeat the first col.
   *  worksheet2.repeat_columns(0, 1); // Repeat the first two cols.
   * @endcode
   */
  void repeat_columns(col_num_t first_col, col_num_t last_col);

  /**
   * @brief Set the worksheet to print in black and white
   *
   * Set the option to print the worksheet in black and white:
   *
   * @code
   *  worksheet.print_black_and_white();
   * @endcode
   */
  void print_black_and_white();

  /**
   * @brief Set the scale factor for the printed page.
   *
   * @param scale Print scale of worksheet to be printed.
   *
   * This function sets the scale factor of the printed page. The Scale factor
   * must be in the range `10 <= scale <= 400`:
   *
   * @code
   *  worksheet1.set_print_scale(75);
   *  worksheet2.set_print_scale(400);
   * @endcode
   *
   * The default scale factor is 100. Note, `%set_print_scale()` does
   * not affect the scale of the visible page in Excel. For that you should use
   * `set_zoom()`.
   *
   * @note that although it is valid to use both `%fit_to_pages()` and
   * `%set_print_scale()` on the same worksheet Excel only allows one
   * of these options to be active at a time. The last function call made will
   * set the active option.
   */
  void set_print_scale(uint16_t scale);

  /**
   * @brief Set the start/first page number when printing.
   *
   * @param start_page Page number of the starting page when printing.
   *
   * The `%set_start_page()` function is used to set the number number
   * of the first page when the worksheet is printed out. It is the same as the
   * "First Page Number" option in Excel:
   *
   * @code
   *  // Start print from page 2.
   *  worksheet.set_start_page(2);
   * @endcode
   */
  void set_start_page(uint16_t start_page);

  // TODO Only used by workbook, should be private.
  void set_error_cell(const object_properties_t& object_props, uint32_t ref_id);

  // TODO Only used by packager, should be private.
  [[nodiscard]] std::string assemble_xml_file();

  static const size_t MAX_NUMBER_URLS = 65530;
  static const row_num_t ROW_MAX      = 1048576;
  static const col_num_t COL_MAX      = 16384;
  static const size_t STR_MAX         = 32767;
  static const col_num_t COL_META_MAX = 128;

private:
  // TODO To be reworked
  friend class packager_t;
  friend class workbook_t;
  friend class rich_value_t;
  friend class chartsheet_t;

  // Retrieve functions
  [[nodiscard]] std::string get_sheet_name() const;
  [[nodiscard]] uint16_t get_sheet_index() const;
  [[nodiscard]] row_t& get_row(row_num_t row_num);

  // Insertion functions.
  void check_dimensions(row_num_t row_num, col_num_t col_num, bool ignore_row, bool ignore_col);
  void insert_cell(row_num_t row_num, col_num_t col_num, const cell_t& cell);
  void insert_hyperlink(row_num_t row_num, col_num_t col_num, const cell_t& link);
  void insert_comment(row_num_t row_num, col_num_t col_num, const cell_t& link);
  void insert_cell_placeholder(row_num_t row_num, col_num_t col_num);
  void store_array_formula(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                           const std::string& formula, const format_t* format, double result, bool is_dynamic);
  void store_conditional_format_object(const cond_format_obj_t& cond_format);
  void validate_conditional_data_bar(cond_format_obj_t& cond_format, const conditional_format_t& user_options);
  void write_table_column_data(const table_obj_t& table_obj);
  void write_column_function(row_num_t row_num, col_num_t col_num, const table_column_t& column);
  void write_column_formula(row_num_t first_row, row_num_t last_row, col_num_t col, const table_column_t& column);

  // Function to prepare data before packaging.
  uint32_t calculate_x_split_width(double x_split) const;
  void set_header_footer_image(const std::string& filename, image_position_t image_position);
  [[nodiscard]] uint32_t prepare_vml_objects(uint32_t vml_data_id, uint32_t vml_shape_id, uint32_t vml_drawing_id,
                                             uint32_t comment_id);
  [[nodiscard]] uint32_t size_col(col_num_t col_num, object_position_t anchor);
  [[nodiscard]] uint32_t size_row(row_num_t row_num, object_position_t anchor);
  [[nodiscard]] const row_t* find_row(row_num_t row_num) const;
  [[nodiscard]] const cell_t* find_cell_in_row(const row_t* row, col_num_t col_num) const;
  void position_object_emus(const object_properties_t& image, drawing_object_t& drawing_object);
  void position_object_pixels(const object_properties_t& object_props, drawing_object_t& drawing_object);
  void position_vml_object(vml_obj_t& vml_obj);
  [[nodiscard]] uint32_t find_drawing_rel_index(const std::string& target);
  [[nodiscard]] std::string get_vml_image_position(image_position_t image_position) const;
  [[nodiscard]] uint32_t find_vml_drawing_rel_index(const std::string& target);
  [[nodiscard]] uint32_t get_drawing_rel_index(const std::string& target);
  [[nodiscard]] uint32_t get_vml_drawing_rel_index(const std::string& target);
  void prepare_image(uint32_t image_ref_id, uint32_t drawing_id, object_properties_t& object_props);
  void prepare_header_image(uint32_t image_ref_id, object_properties_t& object_props);
  void prepare_header_vml_objects(uint32_t vml_header_id, uint32_t vml_drawing_id);
  void prepare_chart(uint32_t chart_ref_id, uint32_t drawing_id, object_properties_t& object_props, bool is_chartsheet);
  void prepare_background(uint32_t image_ref_id, object_properties_t& object_props);
  void prepare_tables(uint32_t table_id);

  // Functions to write XML content.
  [[nodiscard]] std::string write_worksheet() const;
  [[nodiscard]] std::string write_sheet_pr() const;
  [[nodiscard]] std::string write_dimension() const;
  [[nodiscard]] std::string write_sheet_view();
  [[nodiscard]] std::string write_sheet_views();
  [[nodiscard]] std::string write_sheet_format_pr() const;
  [[nodiscard]] std::string write_cols() const;
  [[nodiscard]] std::string write_col_info(const col_options_t& options) const;
  [[nodiscard]] std::string write_sheet_data() const;
  [[nodiscard]] std::string write_auto_filter() const;
  [[nodiscard]] std::string write_merge_cells() const;
  [[nodiscard]] std::string write_conditional_formats();
  [[nodiscard]] std::string write_data_validations() const;
  [[nodiscard]] std::string write_hyperlinks();
  [[nodiscard]] std::string write_print_options() const;
  [[nodiscard]] std::string write_page_margins() const;
  [[nodiscard]] std::string write_page_setup() const;
  [[nodiscard]] std::string write_header_footer() const;
  [[nodiscard]] std::string write_row_breaks() const;
  [[nodiscard]] std::string write_col_breaks() const;
  [[nodiscard]] std::string write_ignored_errors() const;
  [[nodiscard]] std::string write_drawings();
  [[nodiscard]] std::string write_legacy_drawing();
  [[nodiscard]] std::string write_legacy_drawing_hf();
  [[nodiscard]] std::string write_picture();
  [[nodiscard]] std::string write_table_part(uint16_t id);
  [[nodiscard]] std::string write_table_parts();
  [[nodiscard]] std::string write_ext_list();
  [[nodiscard]] std::string write_row(const row_t& row, const std::string& spans) const;
  [[nodiscard]] std::string write_row(const row_t& row) const;
  [[nodiscard]] std::string write_rows() const;
  [[nodiscard]] std::string write_string_cell(std::string_view range, int32_t style_index, const cell_t& cell) const;
  [[nodiscard]] std::string write_number_cell(std::string_view range, int32_t style_index, const cell_t& cell) const;
  [[nodiscard]] std::string write_cell(const cell_t& cell, format_t* row_format) const;
  [[nodiscard]] std::string write_hyperlink_internal(row_num_t row_num, col_num_t col_num, const std::string& location,
                                                     const std::string& display, const std::string& tooltip) const;
  [[nodiscard]] std::string write_hyperlink_external(row_num_t row_num, col_num_t col_num, const std::string& location,
                                                     const std::string& tooltip, uint16_t id) const;
  [[nodiscard]] std::string write_drawing(uint16_t id) const;
  [[nodiscard]] std::string write_odd_header() const;
  [[nodiscard]] std::string write_odd_footer() const;
  [[nodiscard]] std::string write_brk(uint32_t id, uint32_t max) const;
  [[nodiscard]] std::string write_tab_color() const;
  [[nodiscard]] std::string write_merge_cell(const merged_range_t& merged_range) const;
  [[nodiscard]] std::string write_formula_num_cell(const cell_t& cell) const;
  [[nodiscard]] std::string write_formula_str_cell(const cell_t& cell) const;
  [[nodiscard]] std::string write_boolean_cell(const cell_t& cell) const;
  [[nodiscard]] std::string write_error_cell() const;
  [[nodiscard]] std::string write_sheet_protection(const protection_obj_t& protection) const;
  [[nodiscard]] std::string write_filter_column(const std::optional<filter_rule_obj_t>& filter) const;
  [[nodiscard]] std::string write_filter(const std::string& str, double num, filter_criteria_t criteria) const;
  [[nodiscard]] std::string write_filter_standard(const filter_rule_obj_t& filter) const;
  [[nodiscard]] std::string write_custom_filter(const std::string& str, double num, filter_criteria_t criteria) const;
  [[nodiscard]] std::string write_filter_list(const filter_rule_obj_t& filter) const;
  [[nodiscard]] std::string write_filter_custom(const filter_rule_obj_t& filter) const;
  [[nodiscard]] std::string write_array_formula_num_cell(const cell_t& cell) const;
  [[nodiscard]] std::string write_conditional_formatting(const std::string& sqref,
                                                         std::vector<cond_format_obj_t>& cond_formats);
  [[nodiscard]] std::string write_formula_str(const std::string& data) const;
  [[nodiscard]] std::string write_formula_num(double num) const;
  [[nodiscard]] std::string write_color(color_t color) const;
  [[nodiscard]] std::string write_cfvo_str(conditional_format_rule_types_t rule_type, const std::string& value,
                                           bool data_bar_2010) const;
  [[nodiscard]] std::string write_cfvo_num(conditional_format_rule_types_t rule_type, double value,
                                           bool data_bar_2010) const;
  [[nodiscard]] std::string write_cf_rule(cond_format_obj_t& cond_format);
  [[nodiscard]] std::string write_cf_rule_cell(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_text(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_blanks(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_formula(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_color_scale(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_data_bar(cond_format_obj_t& cond_format);
  [[nodiscard]] std::string write_data_bar(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_data_bar_ext(cond_format_obj_t& cond_format);
  [[nodiscard]] std::string write_cf_rule_time_period(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_duplicate(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_average(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_top(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_cf_rule_icons(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_icon_set(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_ext(const std::string& uri) const;
  [[nodiscard]] std::string write_ext_list_data_bars();
  [[nodiscard]] std::string write_conditional_formatting_2010(std::vector<cond_format_obj_t>& cond_formats) const;
  [[nodiscard]] std::string write_x14_cf_rule(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_xm_sqref(const cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_x14_data_bar(cond_format_obj_t& cond_format) const;
  [[nodiscard]] std::string write_x14_cfvo(conditional_format_rule_types_t rule_type, double number,
                                           const std::string& str) const;
  [[nodiscard]] std::string write_x14_color(const std::string& type, color_t color) const;
  [[nodiscard]] std::string write_data_validation(const data_val_obj_t& validation) const;
  [[nodiscard]] std::string write_formula1_num(double number) const;
  [[nodiscard]] std::string write_formula2_num(double number) const;
  [[nodiscard]] std::string write_formula1_str(const std::string& str) const;
  [[nodiscard]] std::string write_formula2_str(const std::string& str) const;
  [[nodiscard]] std::string write_ignored_error(const std::string& ignore_error, const std::string& range) const;
  [[nodiscard]] std::string write_panes();
  [[nodiscard]] std::string write_freeze_panes();
  [[nodiscard]] std::string write_split_panes();
  [[nodiscard]] std::string write_selection(const selection_t& selection) const;
  [[nodiscard]] std::string write_selections() const;
  [[nodiscard]] std::string write_page_set_up_pr() const;
  [[nodiscard]] std::string write_outline_pr() const;

  std::function<int32_t(format_t*)> get_xf_index_;
  std::function<int32_t(format_t*)> get_dxf_index_;
  table_rows_t table_;
  table_rows_t hyperlinks_;
  table_rows_t comments_;
  std::vector<merged_range_t> merged_ranges_;
  std::list<selection_t> selections_;
  std::vector<data_val_obj_t> data_validations_;
  std::map<std::string, std::vector<cond_format_obj_t>> conditional_formats_;
  std::vector<object_properties_t> image_props_;
  std::vector<object_properties_t> embedded_image_props_;
  std::vector<object_properties_t> chart_data_;
  std::map<std::string, uint32_t> drawing_rel_ids_;
  std::map<std::string, uint32_t> vml_drawing_rel_ids_;
  std::vector<vml_obj_t> comment_objs_;
  std::vector<vml_obj_t> header_image_objs_;
  std::vector<vml_obj_t> button_objs_;
  std::vector<table_obj_t> table_objs_;
  row_num_t dim_rowmin_  = ROW_MAX;
  row_num_t dim_rowmax_  = 0;
  col_num_t dim_colmin_  = COL_MAX;
  col_num_t dim_colmax_  = 0;
  shared_strings_t* sst_ = nullptr;
  std::string name_;
  std::string quoted_name_;
  uint16_t index_         = 0;
  bool active_            = true; // TODO Set to true for test, to be removed
  bool selected_          = false;
  bool hidden_            = false;
  uint16_t* active_sheet_ = nullptr;
  uint16_t* first_sheet_  = nullptr;
  bool is_chartsheet_     = false;
  std::vector<col_options_t> col_options_;
  std::vector<format_t*> col_formats_;
  bool col_size_changed_             = false;
  bool row_size_changed_             = false;
  uint16_t fit_height_               = 0;
  uint16_t fit_width_                = 0;
  uint16_t horizontal_dpi_           = 0;
  uint16_t hlink_count_              = 0;
  uint16_t page_start_               = 0;
  uint16_t print_scale_              = 100;
  uint16_t rel_count_                = 0;
  uint16_t vertical_dpi_             = 0;
  uint16_t zoom_                     = 100;
  bool filter_on_                    = false;
  bool fit_page_                     = false;
  bool hcenter_                      = false;
  drawing_orientation_t orientation_ = drawing_orientation_t::PORTRAIT;
  bool outline_changed_              = false;
  bool outline_on_                   = true;
  bool outline_style_                = true;
  bool outline_below_                = true;
  bool outline_right_                = false;
  uint8_t page_order_                = 0;
  bool page_setup_changed_           = false;
  bool page_view_                    = false;
  uint8_t paper_size_                = 0;
  bool print_gridlines_              = false;
  bool print_headers_                = false;
  bool print_options_changed_        = false;
  bool right_to_left_                = false;
  bool screen_gridlines_             = true;
  bool show_zeros_                   = true;
  bool vcenter_                      = false;
  bool zoom_scale_normal_            = true;
  bool black_white_                  = false;
  bool has_dynamic_functions_        = false;
  std::string vba_codename_;
  color_t tab_color_           = color_t::UNSET;
  double margin_left_          = 0.7;
  double margin_right_         = 0.7;
  double margin_top_           = 0.75;
  double margin_bottom_        = 0.75;
  double margin_header_        = 0.3;
  double margin_footer_        = 0.3;
  double default_row_height_   = DEF_ROW_HEIGHT;
  uint32_t default_row_pixels_ = 20;
  uint32_t default_col_pixels_ = 64;
  bool default_row_zeroed_     = false;
  bool default_row_set_        = false;
  uint8_t outline_row_level_   = 0;
  uint8_t outline_col_level_   = 0;
  bool header_footer_changed_  = false;
  std::string header_;
  std::string footer_;
  repeat_rows_t repeat_rows_;
  repeat_cols_t repeat_cols_;
  print_area_t print_area_;
  autofilter_t autofilter_;
  uint16_t max_url_length_ = 2079;
  std::vector<row_num_t> hbreaks_;
  std::vector<col_num_t> vbreaks_;
  uint32_t drawing_rel_id_     = 0;
  uint32_t vml_drawing_rel_id_ = 0;
  std::vector<std::tuple<std::string, std::string, std::string>> external_hyperlinks_;
  std::vector<std::tuple<std::string, std::string, std::string>> external_drawing_links_;
  std::vector<std::tuple<std::string, std::string, std::string>> drawing_links_;
  std::vector<std::tuple<std::string, std::string, std::string>> vml_drawing_links_;
  std::vector<std::tuple<std::string, std::string, std::string>> external_table_links_;
  panes_t panes_;
  std::string top_left_cell_;
  protection_obj_t protection_;
  std::optional<drawing_t> drawing_;
  format_t* default_url_format_;
  bool has_vml_                = false;
  bool has_comments_           = false;
  bool has_header_vml_         = false;
  bool has_background_image_   = false;
  bool has_buttons_            = false;
  bool storing_embedded_image_ = false;
  std::optional<std::tuple<std::string, std::string, std::string>> external_vml_comment_link_;
  std::optional<std::tuple<std::string, std::string, std::string>> external_comment_link_;
  std::optional<std::tuple<std::string, std::string, std::string>> external_vml_header_link_;
  std::optional<std::tuple<std::string, std::string, std::string>> external_background_link_;
  std::string comment_author_;
  std::string vml_data_id_str_;
  std::string vml_header_id_str_;
  uint32_t vml_shape_id_;
  uint32_t vml_header_id_;
  uint32_t dxf_priority_                     = 0;
  comment_display_t comment_display_default_ = comment_display_t::HIDDEN;
  uint32_t data_bar_2010_index_              = 0;
  bool has_ignore_errors_                    = false;
  std::string ignore_number_stored_as_text_;
  std::string ignore_eval_error_;
  std::string ignore_formula_differs_;
  std::string ignore_formula_range_;
  std::string ignore_formula_unlocked_;
  std::string ignore_empty_cell_reference_;
  std::string ignore_list_data_validation_;
  std::string ignore_calculated_column_;
  std::string ignore_two_digit_text_year_;
  bool use_1904_epoch_    = false;
  uint16_t excel_version_ = 0;
  std::optional<object_properties_t> header_footer_objs_[6];
  std::optional<object_properties_t> background_image_;
  std::vector<std::optional<filter_rule_obj_t>> filter_rules_;
  col_num_t num_filter_rules_ = 0;
};

/// @cond
}
/// @endcond

#endif
