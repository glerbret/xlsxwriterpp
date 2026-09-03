/*
 * chart - A Xlsxwriter++ library for creating Excel XLSX chart files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page chart_page The Chart object
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * @see @ref chart.h for full details of the functionality.
 *
 * @file chart.h
 *
 * @brief Functions related to adding data to and configuring a chart.
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * A Chart object isn't created directly. Instead a chart is created by
 * calling the `workbook_t::add_chart()` function from a Workbook object. For
 * example:
 *
 * @code
 *  #include "xlsxwriterpp.h"
 *
 *  int main()
 *  {
 *    xwpp::workbook_t workbook;
 *    xwpp::worksheet& worksheet = workbook.add_worksheet();
 *
 *    // User function to add data to worksheet, not shown here.
 *    write_worksheet_data(worksheet);
 *
 *    // Create a chart object.
 *    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::COLUMN);
 *
 *    // In the simplest case we just add some value data series.
 *    chart.add_series("", "=Sheet1!$A$1:$A$5");
 *    chart.add_series("", "=Sheet1!$B$1:$B$5");
 *    chart.add_series("", "=Sheet1!$C$1:$C$5");
 *
 *    // Insert the chart into the worksheet
 *    worksheet.insert_chart(CELL("B7"), &chart);
 *
 *    return workbook.save("chart.xlsx");
 *  }
 * @endcode
 *
 * The chart in the worksheet will look like this:
 *
 * @image html chart_simple.png
 *
 * The basic procedure for adding a chart to a worksheet is:
 *
 * 1. Create the chart with `workbook_t::add_chart()`.
 * 2. Add one or more data series to the chart which refers to data in the
 *    workbook using `chart_t::add_series()`.
 * 3. Configure the chart with the other available functions shown below.
 * 4. Insert the chart into a worksheet using `worksheet_t::insert_chart()`.
 *
 * @todo Rework, several free functions should be member functions.
 */

#ifndef XWPP_CHART_H
#define XWPP_CHART_H

#include "xwpp/common.h"
#include "xwpp/format.h"

#include <cstdint>
#include <functional>
#include <list>
#include <string>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief Available chart types.
 *
 * Available chart types.
 */
enum class chart_type_t
{
  /** None. */
  NONE = 0,

  /** Area chart. */
  AREA,

  /** Area chart - stacked. */
  AREA_STACKED,

  /** Area chart - percentage stacked. */
  AREA_STACKED_PERCENT,

  /** Bar chart. */
  BAR,

  /** Bar chart - stacked. */
  BAR_STACKED,

  /** Bar chart - percentage stacked. */
  BAR_STACKED_PERCENT,

  /** Column chart. */
  COLUMN,

  /** Column chart - stacked. */
  COLUMN_STACKED,

  /** Column chart - percentage stacked. */
  COLUMN_STACKED_PERCENT,

  /** Doughnut chart. */
  DOUGHNUT,

  /** Line chart. */
  LINE,

  /** Line chart - stacked. */
  LINE_STACKED,

  /** Line chart - percentage stacked. */
  LINE_STACKED_PERCENT,

  /** Pie chart. */
  PIE,

  /** Scatter chart. */
  SCATTER,

  /** Scatter chart - straight. */
  SCATTER_STRAIGHT,

  /** Scatter chart - straight with markers. */
  SCATTER_STRAIGHT_WITH_MARKERS,

  /** Scatter chart - smooth. */
  SCATTER_SMOOTH,

  /** Scatter chart - smooth with markers. */
  SCATTER_SMOOTH_WITH_MARKERS,

  /** Radar chart. */
  RADAR,

  /** Radar chart - with markers. */
  RADAR_WITH_MARKERS,

  /** Radar chart - filled. */
  RADAR_FILLED
};

/**
 * @brief Chart legend positions.
 *
 * Chart legend positions.
 */
enum class chart_legend_position_t
{
  /** No chart legend. */
  NONE = 0,

  /** Chart legend positioned at right side. */
  RIGHT,

  /** Chart legend positioned at left side. */
  LEFT,

  /** Chart legend positioned at top. */
  TOP,

  /** Chart legend positioned at bottom. */
  BOTTOM,

  /** Chart legend positioned at top right. */
  TOP_RIGHT,

  /** Chart legend overlaid at right side. */
  OVERLAY_RIGHT,

  /** Chart legend overlaid at left side. */
  OVERLAY_LEFT,

  /** Chart legend overlaid at top right. */
  OVERLAY_TOP_RIGHT
};

/**
 * @brief Chart line dash types.
 *
 * Chart line dash types.
 *
 * The dash types are shown in the order that they appear in the Excel dialog.
 *
 * @see @ref chart_lines.
 */
enum class chart_line_dash_type_t
{
  /** Solid. */
  SOLID = 0,

  /** Round Dot. */
  ROUND_DOT,

  /** Square Dot. */
  SQUARE_DOT,

  /** Dash. */
  DASH,

  /** Dash Dot. */
  DASH_DOT,

  /** Long Dash. */
  LONG_DASH,

  /** Long Dash Dot. */
  LONG_DASH_DOT,

  /** Long Dash Dot Dot. */
  LONG_DASH_DOT_DOT,

  /* These aren't available in the dialog but are used by Excel. */
  DOT,
  SYSTEM_DASH_DOT,
  SYSTEM_DASH_DOT_DOT
};

/**
 * @brief Chart marker types.
 *
 * Chart marker types.
 */
enum class chart_marker_type_t
{
  /** Automatic, series default, marker type. */
  AUTOMATIC,

  /** No marker type. */
  NONE,

  /** Square marker type. */
  SQUARE,

  /** Diamond marker type. */
  DIAMOND,

  /** Triangle marker type. */
  TRIANGLE,

  /** X shape marker type. */
  X,

  /** Star marker type. */
  STAR,

  /** Short dash marker type. */
  SHORT_DASH,

  /** Long dash marker type. */
  LONG_DASH,

  /** Circle marker type. */
  CIRCLE,

  /** Plus (+) marker type. */
  PLUS
};

/**
 * @brief Chart pattern types.
 *
 * Chart pattern types.
 */
enum class chart_pattern_type_t
{
  /** None pattern. */
  NONE,

  /** 5 Percent pattern. */
  PERCENT_5,

  /** 10 Percent pattern. */
  PERCENT_10,

  /** 20 Percent pattern. */
  PERCENT_20,

  /** 25 Percent pattern. */
  PERCENT_25,

  /** 30 Percent pattern. */
  PERCENT_30,

  /** 40 Percent pattern. */
  PERCENT_40,

  /** 50 Percent pattern. */
  PERCENT_50,

  /** 60 Percent pattern. */
  PERCENT_60,

  /** 70 Percent pattern. */
  PERCENT_70,

  /** 75 Percent pattern. */
  PERCENT_75,

  /** 80 Percent pattern. */
  PERCENT_80,

  /** 90 Percent pattern. */
  PERCENT_90,

  /** Light downward diagonal pattern. */
  LIGHT_DOWNWARD_DIAGONAL,

  /** Light upward diagonal pattern. */
  LIGHT_UPWARD_DIAGONAL,

  /** Dark downward diagonal pattern. */
  DARK_DOWNWARD_DIAGONAL,

  /** Dark upward diagonal pattern. */
  DARK_UPWARD_DIAGONAL,

  /** Wide downward diagonal pattern. */
  WIDE_DOWNWARD_DIAGONAL,

  /** Wide upward diagonal pattern. */
  WIDE_UPWARD_DIAGONAL,

  /** Light vertical pattern. */
  LIGHT_VERTICAL,

  /** Light horizontal pattern. */
  LIGHT_HORIZONTAL,

  /** Narrow vertical pattern. */
  NARROW_VERTICAL,

  /** Narrow horizontal pattern. */
  NARROW_HORIZONTAL,

  /** Dark vertical pattern. */
  DARK_VERTICAL,

  /** Dark horizontal pattern. */
  DARK_HORIZONTAL,

  /** Dashed downward diagonal pattern. */
  DASHED_DOWNWARD_DIAGONAL,

  /** Dashed upward diagonal pattern. */
  DASHED_UPWARD_DIAGONAL,

  /** Dashed horizontal pattern. */
  DASHED_HORIZONTAL,

  /** Dashed vertical pattern. */
  DASHED_VERTICAL,

  /** Small confetti pattern. */
  SMALL_CONFETTI,

  /** Large confetti pattern. */
  LARGE_CONFETTI,

  /** Zigzag pattern. */
  ZIGZAG,

  /** Wave pattern. */
  WAVE,

  /** Diagonal brick pattern. */
  DIAGONAL_BRICK,

  /** Horizontal brick pattern. */
  HORIZONTAL_BRICK,

  /** Weave pattern. */
  WEAVE,

  /** Plaid pattern. */
  PLAID,

  /** Divot pattern. */
  DIVOT,

  /** Dotted grid pattern. */
  DOTTED_GRID,

  /** Dotted diamond pattern. */
  DOTTED_DIAMOND,

  /** Shingle pattern. */
  SHINGLE,

  /** Trellis pattern. */
  TRELLIS,

  /** Sphere pattern. */
  SPHERE,

  /** Small grid pattern. */
  SMALL_GRID,

  /** Large grid pattern. */
  LARGE_GRID,

  /** Small check pattern. */
  SMALL_CHECK,

  /** Large check pattern. */
  LARGE_CHECK,

  /** Outlined diamond pattern. */
  OUTLINED_DIAMOND,

  /** Solid diamond pattern. */
  SOLID_DIAMOND
};

/**
 * @brief Chart data label positions.
 *
 * Chart data label positions.
 */
enum class chart_label_position_t
{
  /** Series data label position: default position. */
  DEFAULT,

  /** Series data label position: center. */
  CENTER,

  /** Series data label position: right. */
  RIGHT,

  /** Series data label position: left. */
  LEFT,

  /** Series data label position: above. */
  ABOVE,

  /** Series data label position: below. */
  BELOW,

  /** Series data label position: inside base.  */
  INSIDE_BASE,

  /** Series data label position: inside end. */
  INSIDE_END,

  /** Series data label position: outside end. */
  OUTSIDE_END,

  /** Series data label position: best fit. */
  BEST_FIT
};

/**
 * @brief Chart data label separator.
 *
 * Chart data label separator.
 */
enum class chart_label_separator_t
{
  /** Series data label separator: comma (the default). */
  COMMA,

  /** Series data label separator: semicolon. */
  SEMICOLON,

  /** Series data label separator: period. */
  PERIOD,

  /** Series data label separator: newline. */
  NEWLINE,

  /** Series data label separator: space. */
  SPACE
};

/**
 * @brief Chart axis types.
 *
 * Chart axis types.
 */
enum class chart_axis_type_t
{
  /** Chart X axis. */
  TYPE_X,

  /** Chart Y axis. */
  TYPE_Y
};

enum class chart_subtype_t
{
  NONE = 0,
  STACKED,
  STACKED_PERCENT
};

enum class chart_grouping_t
{
  CLUSTERED,
  STANDARD,
  PERCENTSTACKED,
  STACKED
};

/**
 * @brief Axis positions for category axes.
 *
 * Axis positions for category axes.
 */
enum class chart_axis_tick_position_t
{
  DEFAULT,

  /** Position category axis on tick marks. */
  ON_TICK,

  /** Position category axis between tick marks. */
  BETWEEN
};

/**
 * @brief Axis label positions.
 *
 * Axis label positions.
 */
enum class chart_axis_label_position_t
{
  /** Position the axis labels next to the axis. The default. */
  NEXT_TO,

  /** Position the axis labels at the top of the chart, for horizontal
   * axes, or to the right for vertical axes.*/
  HIGH,

  /** Position the axis labels at the bottom of the chart, for horizontal
   * axes, or to the left for vertical axes.*/
  LOW,

  /** Turn off the the axis labels. */
  NONE
};

/**
 * @brief Axis label alignments.
 *
 * Axis label alignments.
 */
enum class chart_axis_label_alignment_t
{
  /** Chart axis label alignment: center. */
  CENTER,

  /** Chart axis label alignment: left. */
  LEFT,

  /** Chart axis label alignment: right. */
  RIGHT
};

/**
 * @brief Display units for chart value axis.
 *
 * Display units for chart value axis.
 */
enum class chart_axis_display_unit_t
{
  /** Axis display units: None. The default. */
  NONE,

  /** Axis display units: Hundreds. */
  HUNDREDS,

  /** Axis display units: Thousands. */
  THOUSANDS,

  /** Axis display units: Ten thousands. */
  TEN_THOUSANDS,

  /** Axis display units: Hundred thousands. */
  HUNDRED_THOUSANDS,

  /** Axis display units: Millions. */
  MILLIONS,

  /** Axis display units: Ten millions. */
  TEN_MILLIONS,

  /** Axis display units: Hundred millions. */
  HUNDRED_MILLIONS,

  /** Axis display units: Billions. */
  BILLIONS,

  /** Axis display units: Trillions. */
  TRILLIONS
};

/**
 * @brief Tick mark types for an axis.
 *
 * Tick mark types for an axis.
 */
enum class chart_axis_tick_mark_t
{
  /** Default tick mark for the chart axis. Usually outside. */
  DEFAULT,

  /** No tick mark for the axis. */
  NONE,

  /** Tick mark inside the axis only. */
  INSIDE,

  /** Tick mark outside the axis only. */
  OUTSIDE,

  /** Tick mark inside and outside the axis. */
  CROSSING
};

struct series_data_point_t
{
  bool is_string_{false};
  double number_{0};
  std::string str_;
  bool no_data_{false};
};

// TODO struct or class
struct series_range_t
{
  std::string formula_;
  std::string sheetname_;
  row_num_t first_row_{0};
  row_num_t last_row_{0};
  col_num_t first_col_{0};
  col_num_t last_col_{0};
  bool ignore_cache_{false};

  bool has_string_cache_{false};
  uint16_t num_data_points_{0};
  std::vector<series_data_point_t> data_cache_;
};

/**
 * @brief Struct to represent a chart line.
 *
 * Struct to represent a chart line.
 *
 * See @ref chart_lines.
 */
struct chart_line_t
{
  /**
   * The chart font color.
   *
   * @see @ref working_with_colors.
   */
  color_t color_{color_t::UNSET};

  /** Turn off/hide line.*/
  bool none_{false};

  /** Width of the line in increments of 0.25. Default is 2.25. */
  double width_{0.};

  /**
   * The line dash type.
   *
   * @see `chart_line_dash_type_t`.
   */
  chart_line_dash_type_t dash_type_{chart_line_dash_type_t::SOLID};

  /** Set the transparency of the line. 0 - 100. Default 0. */
  uint8_t transparency_{0};
};

/**
 * @brief Struct to represent a chart fill.
 *
 * Struct to represent a chart fill.
 *
 * @see @ref chart_fills.
 */
struct chart_fill_t
{
  /**
   * The chart font color.
   *
   * @see @ref working_with_colors.
   */
  color_t color_{color_t::UNSET};

  /** Turn off/hide line.*/
  bool none_{false};

  /** Set the transparency of the fill. 0 - 100. Default 0. */
  uint8_t transparency_{0};
};

/**
 * @brief Struct to represent a chart pattern.
 *
 * Struct to represent a chart pattern.
 *
 * @see @ref chart_patterns.
 */
struct chart_pattern_t
{
  /**
   * The pattern foreground color.
   *
   * @see @ref working_with_colors.
   */
  color_t fg_color_{color_t::UNSET};

  /**
   * The pattern background color.
   *
   * @see @ref working_with_colors.
   */
  color_t bg_color_{color_t::UNSET};

  /** The pattern type. See `chart_pattern_type_t`. */
  chart_pattern_type_t type_{chart_pattern_type_t::NONE};
};

/**
 * @brief Struct to represent a chart font.
 *
 * Struct to represent a chart font.
 *
 * @see @ref chart_fonts.
 */
struct chart_font_t
{
  /** The chart font name, such as `"Arial"` or `"Calibri"`. */
  std::string name_;

  /** The chart font size. The default is 11. */
  double size_{0.};

  /** The chart font bold property. Set to 0 or 1. */
  // By default, bold_ may be set to true (title) or false.
  // So with use optional to indicate unset value and get default behaviour.
  std::optional<bool> bold_;

  /** The chart font italic property. Set to 0 or 1. */
  bool italic_{false};

  /** The chart font underline property. Set to 0 or 1. */
  bool underline_{false};

  /**
   * The chart font rotation property. Range: -90 to 90, and 270, 271 and 360:
   *
   *  - The angles -90 to 90 are the normal range shown in the Excel user interface.
   *  - The angle 270 gives a stacked (top to bottom) alignment.
   *  - The angle 271 gives a stacked alignment for East Asian fonts.
   *  - The angle 360 gives an explicit angle of 0 to override the y axis default.
   */
  int32_t rotation_{0};

  /**
   * The chart font color.
   *
   * @see @ref working_with_colors.
   */
  color_t color_{color_t::UNSET};

  /** The chart font pitch family property. Rarely required, set to 0. */
  uint8_t pitch_family_{0};

  /** The chart font character set property. Rarely required, set to 0. */
  uint8_t charset_{0};

  /** The chart font baseline property. Rarely required, set to 0. */
  int8_t baseline_{0};

  // Use to distinguish title font (bold by default and we should set b="0" if bold_ is false) from others one
  bool title_font_{false};
};

/**
 * @brief Struct to represent Excel chart element layout dimensions.
 *
 * Excel supports manual positioning of elements such as the chart axis labels,
 * the chart legend, the chart plot area and the chart title. The
 * `%chart_layout_t` struct represents the layout dimension for these elements.
 *
 * The layout units used by Excel are relative units expressed as a percentage
 * of the chart dimensions and are double values in the range `0.0 < x <= 1.0`.
 * Excel calculates these dimensions as shown below:
 *
 * @image html chart_layout.png
 *
 * With reference to the above figure the layout units are calculated as
 * follows:
 *
 * ```text
 *     x = a / W
 *     y = b / H
 * ```
 *
 * These units are cumbersome and can vary depending on other elements in the
 * chart such as text lengths. However, these are the units that are required by
 * Excel to allow relative positioning. Some trial and error is generally
 * required.
 *
 * For the chart `chart_t::plotarea_set_layout()` and `chart_t::legend_set_layout()`
 * functions you can also set the width and height based on the following
 * calculation:
 *
 * ```text
 *     width  = w / W
 *     height = h / H
 * ```
 *
 * For other text based objects the width and height are changed by the font
 * dimensions.
 *
 * The chart functions that support `%chart_layout_t` are:
 *
 * - `chart_t::title_set_layout()`
 * - `chart_t::legend_set_layout()`
 * - `chart_t::plotarea_set_layout()`
 * - `chart_axis_set_name_layout()`
 *
 */
struct chart_layout_t
{
  /** The x offset in the range `0.0 < x <= 1.0` */
  double x_{0.};

  /** The y offset in the range `0.0 < y <= 1.0` */
  double y_{0.};

  /** The width of the plotarea or legend in the range `0.0 < x <= 1.0` */
  double width_{0.};

  /** The height of the plotarea or legend in the range `0.0 < x <= 1.0` */
  double height_{0.};

  bool has_inner_{false};
};

struct chart_marker_t
{
  chart_marker_type_t type_{chart_marker_type_t::AUTOMATIC};
  uint8_t size_{0};
  std::optional<chart_line_t> line_;
  std::optional<chart_fill_t> fill_;
  std::optional<chart_pattern_t> pattern_;
};

struct chart_legend_t
{
  std::optional<chart_font_t> font_;
  chart_legend_position_t position_{chart_legend_position_t::NONE};
  std::optional<chart_layout_t> layout_;
};

struct chart_title_t
{
  std::string name_;
  // TODO ?   lxw_row_t row;
  // TODO ?   lxw_col_t col;
  std::optional<chart_font_t> font_;
  bool off_{false};
  bool is_horizontal_{false};
  // TODO ?   uint8_t ignore_cache;
  bool has_overlay_{false};

  // TODO ?
  /* We use a range to hold the title formula properties even though it
   * will only have 1 point in order to re-use similar functions.*/
  series_range_t range_;

  // TODO ?   struct lxw_series_data_point data_point;
  std::optional<chart_layout_t> layout_;
};

/**
 * @brief Struct to represent an Excel chart data point.
 *
 * The `%chart_point` used to set the line, fill and pattern of one or more
 * points in a chart data series.
 *
 * @see @ref chart_points.
 */
struct chart_point_t
{
  /**
   * The line/border for the chart point.
   *
   * @see @ref chart_lines.
   */
  std::optional<chart_line_t> line_;

  /**
   * The fill for the chart point.
   *
   * @see @ref chart_fills.
   */
  std::optional<chart_fill_t> fill_;

  /**
   * The pattern for the chart point.
   *
   * @see @ref chart_patterns.
   */
  std::optional<chart_pattern_t> pattern_;
};

/**
 * @brief Struct to represent an Excel chart data label.
 *
 * The `%chart_data_label_t` struct is used to represent a data label in a
 * chart series so that custom properties can be set for it.
 */
struct chart_data_label_t
{
  /**
   * The string or formula value for the data label.
   *
   * @see @ref chart_custom_labels.
   */
  std::string value_;

  /**
   * Option to hide/delete the data label from the chart series.
   *
   * @see @ref chart_custom_labels.
   */
  bool hide_{false};

  /**
   * The font properties for the chart data label.
   *
   * @see @ref chart_fonts.
   */
  std::optional<chart_font_t> font_;

  /**
   * The line/border for the chart data label.
   *
   * @see @ref chart_lines.
   */
  std::optional<chart_line_t> line_;

  /**
   * The fill for the chart data label.
   *
   * @see @ref chart_fills.
   */
  std::optional<chart_fill_t> fill_;

  /**
   * The pattern for the chart data label.
   *
   * @see @ref chart_patterns.
   */
  std::optional<chart_pattern_t> pattern_;
};

// Internal version of lxw_chart_data_label with more metadata.
struct chart_custom_label_t
{
  std::string value_;
  bool hide_{false};
  std::optional<chart_font_t> font_;
  std::optional<chart_line_t> line_;
  std::optional<chart_fill_t> fill_;
  std::optional<chart_pattern_t> pattern_;

  // We use a range to hold the label formula properties even though it
  // will only have 1 point in order to re-use similar functions.
  std::optional<series_range_t> range_;

  // TODO ?   struct lxw_series_data_point data_point;
};

/**
 * @brief Define how blank values are displayed in a chart.
 *
 * Define how blank values are displayed in a chart.
 */
enum class chart_blank_t
{
  /** Show empty chart cells as gaps in the data. The default. */
  AS_GAP,

  /** Show empty chart cells as zeros. */
  AS_ZERO,

  /** Show empty chart cells as connected. Only for charts with lines. */
  AS_CONNECTED
};

enum class chart_position_t
{
  RIGHT,
  LEFT,
  TOP,
  BOTTOM
};

enum class chart_layout_type_t
{
  TITLE,
  LEGEND,
  PLOTAREA,
  AXIS_NAME
};

/**
 * @brief Type/amount of data series error bar.
 *
 * Type/amount of data series error bar.
 */
enum class chart_error_bar_type_t
{
  /** Error bar type: Standard error. */
  STD_ERROR,

  /** Error bar type: Fixed value. */
  FIXED,

  /** Error bar type: Percentage. */
  PERCENTAGE,

  /** Error bar type: Standard deviation(s). */
  STD_DEV
};

/**
 * @brief Direction for a data series error bar.
 *
 * Direction for a data series error bar.
 */
enum class chart_error_bar_direction_t
{

  /** Error bar extends in both directions. The default. */
  BOTH,

  /** Error bar extends in positive direction. */
  PLUS,

  /** Error bar extends in negative direction. */
  MINUS
};

/**
 * @brief Direction for a data series error bar.
 *
 * Direction for a data series error bar.
 */
enum class chart_error_bar_axis_t
{
  /** X axis error bar. */
  X,

  /** Y axis error bar. */
  Y
};

/**
 * @brief End cap styles for a data series error bar.
 *
 * End cap styles for a data series error bar.
 */
enum class chart_error_bar_cap_t
{
  /** Flat end cap. The default. */
  END_CAP,

  /** No end cap. */
  NO_CAP
};

struct series_error_bars_t
{
  chart_error_bar_type_t type_{chart_error_bar_type_t::STD_ERROR};
  chart_error_bar_direction_t direction_{chart_error_bar_direction_t::BOTH};
  chart_error_bar_cap_t endcap_{chart_error_bar_cap_t::END_CAP};
  bool has_value_{false};
  bool is_set_{false};
  bool is_x_{false};
  chart_type_t chart_group_{chart_type_t::NONE};
  double value_{0.};
  std::optional<chart_line_t> line_;
};

/**
 * @brief Series trendline/regression types.
 *
 * Series trendline/regression types.
 */
enum class chart_trendline_type_t
{
  /** Trendline type: Linear. */
  LINEAR,

  /** Trendline type: Logarithm. */
  LOG,

  /** Trendline type: Polynomial. */
  POLY,

  /** Trendline type: Power. */
  POWER,

  /** Trendline type: Exponential. */
  EXP,

  /** Trendline type: Moving Average. */
  AVERAGE
};

/**
 * @brief Struct to represent an Excel chart data series.
 *
 * The `%chart_series_t` is created using the chart_add_series function. It is
 * used in functions that modify a chart series but the members of the struct
 * aren't modified directly.
 *
 * @todo Use a class and convert functions to membre functions.
 */
struct chart_series_t
{
  series_range_t categories_;
  series_range_t values_;
  chart_title_t title_;
  std::optional<chart_line_t> line_;
  std::optional<chart_fill_t> fill_;
  std::optional<chart_pattern_t> pattern_;
  std::optional<chart_marker_t> marker_;
  std::vector<chart_point_t> points_;
  std::vector<chart_custom_label_t> data_labels_;

  bool smooth_{false};
  bool invert_if_negative_{false};

  // Data label parameters.
  bool has_labels_{false};
  bool show_labels_value_{false};
  bool show_labels_category_{false};
  bool show_labels_name_{false};
  bool show_labels_leader_{false};
  bool show_labels_legend_{false};
  bool show_labels_percent_{false};
  chart_label_position_t label_position_{chart_label_position_t::DEFAULT};
  chart_label_separator_t label_separator_{chart_label_separator_t::COMMA};
  chart_label_position_t default_label_position_{chart_label_position_t::DEFAULT};
  std::string label_num_format_;
  std::optional<chart_font_t> label_font_;
  std::optional<chart_line_t> label_line_;
  std::optional<chart_fill_t> label_fill_;
  std::optional<chart_pattern_t> label_pattern_;

  series_error_bars_t x_error_bars_;
  series_error_bars_t y_error_bars_;

  bool has_trendline_{false};
  bool has_trendline_forecast_{false};
  bool has_trendline_equation_{false};
  bool has_trendline_r_squared_{false};
  bool has_trendline_intercept_{false};
  chart_trendline_type_t trendline_type_{chart_trendline_type_t::LINEAR};
  uint8_t trendline_value_{0};
  double trendline_forward_{0.};
  double trendline_backward_{0.};
  chart_trendline_type_t trendline_value_type_{chart_trendline_type_t::LINEAR};
  std::string trendline_name_;
  std::optional<chart_line_t> trendline_line_;
  double trendline_intercept_{0.};
};

// Struct for major/minor axis gridlines.
struct chart_gridline_t
{
  bool visible_{false};
  std::optional<chart_line_t> line_;
};

/**
 * @brief Struct to represent an Excel chart axis.
 *
 * The `%chart_axis_t` struct is used in functions that modify a chart axis
 * but the members of the struct aren't modified directly.
 *
 * @todo Use a class and convert functions to membre functions.
 */
struct chart_axis_t
{
  chart_title_t title_;
  std::string num_format_;
  std::string default_num_format_;
  uint8_t source_linked_{0};

  chart_axis_tick_mark_t major_tick_mark_{chart_axis_tick_mark_t::DEFAULT};
  chart_axis_tick_mark_t minor_tick_mark_{chart_axis_tick_mark_t::DEFAULT};
  // TODO ?   uint8_t is_horizontal;

  chart_gridline_t major_gridlines_;
  chart_gridline_t minor_gridlines_;

  std::optional<chart_font_t> num_font_;
  std::optional<chart_line_t> line_;
  std::optional<chart_fill_t> fill_;
  std::optional<chart_pattern_t> pattern_;

  bool is_category_{false};
  // TODO ?   uint8_t is_date;
  bool is_value_{false};

  chart_position_t axis_position_{chart_position_t::RIGHT};
  chart_axis_tick_position_t position_axis_{chart_axis_tick_position_t::DEFAULT};
  chart_axis_label_position_t label_position_{chart_axis_label_position_t::NEXT_TO};
  chart_axis_label_alignment_t label_align_{chart_axis_label_alignment_t::CENTER};
  bool hidden_{false};
  bool reverse_{false};

  bool has_min_{false};
  double min_{0};
  bool has_max_{false};
  double max_{0};

  bool has_major_unit_{false};
  double major_unit_{0.};
  bool has_minor_unit_{false};
  double minor_unit_{0.};

  uint16_t interval_unit_{0};
  uint16_t interval_tick_{0};

  uint16_t log_base_{0};

  chart_axis_display_unit_t display_units_{chart_axis_display_unit_t::NONE};
  bool display_units_visible_{false};

  bool has_crossing_{false};
  bool crossing_min_{false};
  bool crossing_max_{false};
  double crossing_{0.};
};

/**
 * @brief Representation of an Excel chart.
 *
 * Representation of an Excel chart.
 */
class chart_t
{
public:
  // TODO Constructor should not be public but only used by `workbook_t`.
  explicit chart_t(chart_type_t type);

  /**
   * @brief Add a data series to a chart.
   *
   * @param categories The range of categories in the data series.
   * @param values     The range of values in the data series.
   *
   * @return A reference to `chart_series_t`.
   *
   * In Excel a chart **series** is a collection of information that defines
   * which data is plotted such as the categories and values. It is also used to
   * define the formatting for the data.
   *
   * For a `Xlsxwriter++` chart object the `%chart_t::add_series()` function is
   * used to set the categories and values of the series:
   *
   * @code
   *  chart.add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
   * @endcode
   *
   * The series parameters are:
   *
   * - `%categories`: This sets the chart category labels. The category is more
   *   or less the same as the X axis. In most Excel chart types the
   *   `%categories` property is optional and the chart will just assume a
   *   sequential series from `1..n`:
   *
   * @code
   *  // The empty category will default to 1 to 5 like in Excel.
   *  chart.add_series("", "Sheet1!$A$1:$A$5");
   * @endcode
   *
   *  - `%values`: This is the most important property of a series and is the
   *    only mandatory option for every chart object. This parameter links the
   *    chart with the worksheet data that it displays.
   *
   * The `%categories` and `%values` should be a string formula like
   * `"=Sheet1!$A$2:$A$7"` in the same way it is represented in Excel. This is
   * convenient when recreating a chart from an example in Excel but it is
   * trickier to generate programmatically. For these cases you can set the
   * `%categories` and `%values` to an empty string and use the
   * `chart_series_set_categories()` and `chart_series_set_values()` functions:
   *
   * @code
   *  xwpp::chart_series_t& series = chart.add_series("", "");
   *
   *  // Configure the series using a syntax that is easier to define programmatically.
   *  chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
   *  chart_series_set_values(series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
   * @endcode
   *
   * As shown in the previous example the return value from
   * `%add_series()` is a `chart_series_t` reference. This can be used in
   * other functions that configure a series.
   *
   * More than one series can be added to a chart. The series numbering and
   * order in the Excel chart will be the same as the order in which they are
   * added in `Xlsxwriter++`:
   *
   * @code
   *  chart.add_series("", "Sheet1!$A$1:$A$5");
   *  chart.add_series("", "Sheet1!$B$1:$B$5");
   *  chart.add_series("", "Sheet1!$C$1:$C$5");
   * @endcode
   *
   * It is also possible to specify non-contiguous ranges:
   *
   * @code
   *  chart.add_series("=(Sheet1!$A$1:$A$9,Sheet1!$A$14:$A$25)",
   *                   "=(Sheet1!$B$1:$B$9,Sheet1!$B$14:$B$25)");
   * @endcode
   */
  chart_series_t& add_series(const std::string& categories, const std::string& values);

  /**
   * @brief Set the chart style type.
   *
   * @param style_id An index representing the chart style, 1 - 48.
   *
   * The `%set_style()` function is used to set the style of the chart to
   * one of the 48 built-in styles available on the "Design" tab in Excel 2007:
   *
   * @code
   *  chart.set_style(37)
   * @endcode
   *
   * @image html chart_style.png
   *
   * The style index number is counted from 1 on the top left in the Excel
   * dialog. The default style is 2.
   *
   * @note In Excel 2013 the Styles section of the "Design" tab in Excel shows what
   * were referred to as "Layouts" in previous versions of Excel. These layouts
   * are not defined in the file format. They are a collection of modifications
   * to the base chart type. They can not be defined by the `%set_style()`
   * function.
   */
  void set_style(uint8_t style_id);

  /**
   * @brief Set the title of the chart.
   *
   * @param name The chart title name.
   *
   * The `%title_set_name()` function sets the name (title) for the
   * chart. The name is displayed above the chart.
   *
   * @code
   *  chart.title_set_name("Year End Results");
   * @endcode
   *
   * @image html chart_title_set_name.png
   *
   * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
   * a cell in the workbook that contains the name:
   *
   * @code
   *  chart.title_set_name("=Sheet1!$B$1");
   * @endcode
   *
   * The Excel default is to have no chart title.
   *
   * @see the `title_set_name_range()` function to see how to set the
   * name formula programmatically.
   */
  void title_set_name(const std::string& name);

  /**
   * @brief Set the name of a chart series range.
   *
   * @param series A series object created via `add_series()`.
   * @param name   The series name.
   *
   * The `%series_set_name` function is used to set the name for a chart
   * data series. The series name in Excel is displayed in the chart legend and
   * in the formula bar. The name property is optional and if it isn't supplied
   * it will default to `Series 1..n`.
   *
   * The function applies to a `chart_series_t` object created using
   * `add_series()`:
   *
   * @code
   *  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$B$2:$B$7");
   *
   *  chart.series_set_name(series, "Quarterly budget data");
   * @endcode
   *
   * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
   * a cell in the workbook that contains the name:
   *
   * @code
   *  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$B$2:$B$7");
   *
   *  chart.series_set_name(series, "=Sheet1!$B$1");
   * @endcode
   *
   * See also the `chart_series_set_name_range()` function to see how to set the
   * name formula programmatically.
   *
   * @todo Should be member of `chart_series_t`.
   */
  void series_set_name(chart_series_t& series, const std::string& name) const;

  /**
   * @brief Set the font properties for a chart title.
   *
   * @param font A reference to a chart `chart_font_t` font struct.
   *
   * The `%title_set_name_font()` function is used to set the font of a
   * chart title:
   *
   * @code
   *  xwpp::chart_font_t font{.color_ = xwpp::color_t::BLUE};
   *
   *  chart.title_set_name("Year End Results");
   *  chart.title_set_name_font(font);
   * @endcode
   *
   * @image html chart_title_set_name_font.png
   *
   * @image html chart_title_set_name_font2.png
   *
   * @see @ref chart_fonts.
   */
  void title_set_name_font(const chart_font_t& font);

  /**
   * @brief Set the position of the chart legend.
   *
   * @param position The `chart_legend_position_t` value for the legend.
   *
   * The `%legend_set_position()` function is used to set the chart
   * legend to one of the `chart_legend_position_t` values:
   *
   *  - `%chart_legend_position_t::NONE`
   *  - `%chart_legend_position_t::RIGHT`
   *  - `%chart_legend_position_t::LEFT`
   *  - `%chart_legend_position_t::TOP`
   *  - `%chart_legend_position_t::BOTTOM`
   *  - `%chart_legend_position_t::TOP_RIGHT`
   *  - `%chart_legend_position_t::OVERLAY_RIGHT`
   *  - `%chart_legend_position_t::OVERLAY_LEFT`
   *  - `%chart_legend_position_t::OVERLAY_TOP_RIGHT`
   *
   * @code
   *  chart.legend_set_position(xwpp::chart_legend_position_t::::BOTTOM);
   * @endcode
   *
   * @image html chart_legend_bottom.png
   *
   * This function can also be used to turn off a chart legend:
   *
   * @code
   *  chart_legend_set_position(xwpp::chart_legend_position_t::NONE);
   * @endcode
   *
   * @image html chart_legend_none.png
   */
  void legend_set_position(chart_legend_position_t position);

  /**
   * @brief Turn on a data table below the horizontal axis.
   *
   * The `%set_table()` function adds a data table below the horizontal
   * axis with the data used to plot the chart:
   *
   * @code
   *  // Turn on the data table with default options.
   *  chart.set_table();
   * @endcode
   *
   * @image html chart_data_table1.png
   *
   * The data table can only be shown with Bar, Column, Line and Area charts.
   */
  void set_table();

  /**
   * @brief Turn on/off grid options for a chart data table.
   *
   * @param horizontal  Turn on/off the horizontal grid lines in the table.
   * @param vertical    Turn on/off the vertical grid lines in the table.
   * @param outline     Turn on/off the outline lines in the table.
   * @param legend_keys Turn on/off the legend keys in the table.
   *
   * The `%set_table_grid()` function turns on/off grid options for a
   * chart data table. The data table grid options in Excel are shown in the
   * dialog below:
   *
   * @image html chart_data_table3.png
   *
   * These options can be passed to the `%set_table_grid()` function.
   * The values for a default chart are:
   *
   * - `%horizontal`: On.
   * - `%vertical`: On.
   * - `%outline`:  On.
   * - `%legend_keys`: Off.
   *
   * @code
   *  // Turn on the data table with default options.
   *  chart.set_table();
   *
   *  // Turn on all grid lines and the grid legend.
   *  chart.set_table_grid(true, true, true, true);
   *
   *  // Turn off the legend since it is show in the table.
   *  chart.legend_set_position(xwpp::chart_legend_position_t::NONE);
   *
   * @endcode
   *
   * @image html chart_data_table2.png
   *
   * The data table can only be shown with Bar, Column, Line and Area charts.
   */
  void set_table_grid(bool horizontal, bool vertical, bool outline, bool legend_keys);

  /**
   * @brief Turn on and format Drop Lines for a chart.
   *
   * @param line A `chart_line_t` struct.
   *
   * The `%set_drop_lines()` function adds Drop Lines to charts to
   * show the Category value of points in the data:
   *
   * @code
   *  chart.set_drop_lines(std::nullopt);
   * @endcode
   *
   * @image html chart_data_tools6.png
   *
   * It is possible to format the Drop Line line properties if required:
   *
   * @code
   *  xwpp::chart_line_t line{
   *    .color_     = xwpp::color_t::RED,
   *    .dash_type_ = xwpp::chart_line_dash_type_t::SQUARE_DOT,
   *  };
   *
   *  chart.set_drop_lines(line);
   * @endcode
   *
   * Drop Lines are only available in Line and Area charts.
   *
   * @see @ref chart_lines.
   */
  void set_drop_lines(const std::optional<chart_line_t>& line);

  /**
   * @brief Turn on and format high-low Lines for a chart.
   *
   * @param line A `chart_line_t` struct.
   *
   * The `%set_high_low_lines()` function adds High-Low Lines to charts
   * to show the Category value of points in the data:
   *
   * @code
   *  chart.set_high_low_lines(std::nullopt);
   * @endcode
   *
   * @image html chart_data_tools5.png
   *
   * It is possible to format the High-Low Line line properties if required:
   *
   * @code
   *  xwpp::chart_line_t line{
   *    .color_     = xwpp::color_t::RED,
   *    .dash_type_ = xwpp::chart_line_dash_type_t::SQUARE_DOT,
   *  };
   *
   *  chart_set_high_low_lines(line);
   * @endcode
   *
   * High-Low Lines are only available in Line charts.
   *
   * @see @ref chart_lines.
   */
  void set_high_low_lines(const std::optional<chart_line_t>& line);

  /**
   * @brief Turn on up-down bars for the chart.
   *
   * The `%set_up_down_bars()` function adds Up-Down bars to Line charts
   * to indicate the difference between the first and last data series:
   *
   * @code
   *  chart.set_up_down_bars();
   * @endcode
   *
   * @image html chart_data_tools4.png
   *
   * Up-Down bars are only available in Line charts. By default Up-Down bars are
   * black and white like in the above example. To format the border or fill
   * of the bars see the `set_up_down_bars_format()` function below.
   */
  void set_up_down_bars();

  /**
   * @brief Turn on up-down bars for the chart, with formatting.
   *
   * @param up_bar_line   A `chart_line_t` struct for the up-bar border.
   * @param up_bar_fill   A `chart_fill_t` struct for the up-bar fill.
   * @param down_bar_line A `chart_line_t` struct for the down-bar border.
   * @param down_bar_fill A `chart_fill_t` struct for the down-bar fill.
   *
   * The `%set_up_down_bars_format()` function adds Up-Down bars to Line
   * charts to indicate the difference between the first and last data series.
   * It also allows the up and down bars to be formatted:
   *
   * @code
   *  xwpp::chart_line_t line{
   *    .color_ = xwpp::color_t::BLACK};
   *  xwpp::chart_fill_t up_fill{
   *    .color_ = xwpp::color_t(0x00B050)};
   *  xwpp::chart_fill_t down_fill{
   *    .color_ = xwpp::color_t::RED};
   *
   *  chart.set_up_down_bars_format(line, up_fill, line, down_fill);
   * @endcode
   *
   * @image html chart_up_down_bars.png
   *
   * Up-Down bars are only available in Line charts.
   *
   * @see @ref chart_lines and @ref chart_fills.
   */
  void set_up_down_bars_format(const std::optional<chart_line_t>& up_bar_line,
                               const std::optional<chart_fill_t>& up_bar_fill,
                               const std::optional<chart_line_t>& down_bar_line,
                               const std::optional<chart_fill_t>& down_bar_fill);

  /**
   * @brief Set the data marker type for a series.
   *
   * @param series A series object created via `add_series()`.
   * @param type   The marker type, see `chart_marker_type_t`.
   *
   * In Excel a chart marker is used to distinguish data points in a plotted
   * series. In general only Line and Scatter and Radar chart types use
   * markers. The `Xlsxwriter++` chart types that can have markers are:
   *
   * - `chart_type_t::LINE`
   * - `chart_type_t::SCATTER`
   * - `chart_type_t::SCATTER_STRAIGHT`
   * - `chart_type_t::SCATTER_STRAIGHT_WITH_MARKERS`
   * - `chart_type_t::SCATTER_SMOOTH`
   * - `chart_type_t::SCATTER_SMOOTH_WITH_MARKERS`
   * - `chart_type_t::RADAR`
   * - `chart_type_t::RADAR_WITH_MARKERS`
   *
   * The chart types with `MARKERS` in the name have markers with default colors
   * and shapes turned on by default but it is possible using the various
   * `%series_set_marker_xxx()` functions below to change these defaults. It
   * is also possible to turn on an off markers.
   *
   * The `%series_set_marker_type()` function is used to specify the
   * type of the series marker:
   *
   * @code
   *  chart.series_set_marker_type(series, xwpp::chart_marker_type_t::DIAMOND);
   * @endcode
   *
   * @image html chart_marker1.png
   *
   * The available marker types defined by `chart_marker_type_t` are:
   *
   * - `%chart_marker_type_t::AUTOMATIC`
   * - `%chart_marker_type_t::NONE`
   * - `%chart_marker_type_t::SQUARE`
   * - `%chart_marker_type_t::DIAMOND`
   * - `%chart_marker_type_t::TRIANGLE`
   * - `%chart_marker_type_t::X`
   * - `%chart_marker_type_t::STAR`
   * - `%chart_marker_type_t::SHORT_DASH`
   * - `%chart_marker_type_t::LONG_DASH`
   * - `%chart_marker_type_t::CIRCLE`
   * - `%chart_marker_type_t::PLUS`
   *
   * The `%chart_marker_type_t::NONE` type can be used to turn off default markers:
   *
   * @code
   *  chart.series_set_marker_type(series, xwpp::chart_marker_type_t::NONE);
   * @endcode
   *
   * @image html chart_series_set_marker_none.png
   *
   * The `%chart_marker_type_t::AUTOMATIC` type is a special case which turns on a
   * marker using the default marker style for the particular series. If
   * automatic is on then other marker properties such as size, line or fill
   * cannot be set.
   */
  void series_set_marker_type(chart_series_t& series, chart_marker_type_t type) const;

  /**
   * Set the X or Y error bars for a chart series.
   *
   * @param error_bars A pointer to the series X or Y error bars.
   * @param type       The type of error bar: `chart_error_bar_type_t`.
   * @param value      The error value.
   *
   * Error bars can be added to a chart series to indicate error bounds in the
   * data. The error bars can be vertical `%y_error_bars_` (the most common type)
   * or horizontal `%x_error_bars_` (for Bar and Scatter charts only).
   *
   * @image html chart_error_bars0.png
   *
   * The `%series_set_error_bars()` function sets the error bar type
   * and value associated with the type:
   *
   * @code
   *  xwpp::chart_series_t& series = chart.add_series("=Sheet1!$A$1:$A$5",
   *                                                  "=Sheet1!$B$1:$B$5");
   *
   *  chart.series_set_error_bars(series.y_error_bars_,
   *    xwpp::chart_error_bar_type_t::STD_ERROR, 0);
   * @endcode
   *
   * @image html chart_error_bars1.png
   *
   * The error bar types that be used are:
   *
   * - `%chart_error_bar_type_t::STD_ERROR`: Standard error.
   * - `%chart_error_bar_type_t::FIXED`: Fixed value.
   * - `%chart_error_bar_type_t::PERCENTAGE`: Percentage.
   * - `%chart_error_bar_type_t::STD_DEV`: Standard deviation(s).
   *
   * @note Custom error bars are not currently supported.
   *
   * All error bar types, apart from Standard error, should have a valid
   * value to set the error range:
   *
   * @code
   *  chart.series_set_error_bars(series1.y_error_bars,
   *    xwpp::chart_error_bar_type_t::FIXED, 2);
   *
   *  chart.series_set_error_bars(series2.y_error_bars,
   *    xwpp::chart_error_bar_type_t::PERCENTAGE, 5);
   *
   *  chart.series_set_error_bars(series3.y_error_bars,
   *    xwpp::chart_error_bar_type_t::STD_DEV, 1);
   * @endcode
   *
   * For the Standard error type the value is ignored.
   *
   * @see @ref chart_error_bars.
   */
  void series_set_error_bars(series_error_bars_t& error_bars, chart_error_bar_type_t type, double value) const;

  /**
   * @brief Set the Pie/Doughnut chart rotation.
   *
   * @param rotation The angle of rotation.
   *
   * The `%set_rotation()` function is used to set the rotation of the
   * first segment of a Pie/Doughnut chart. This has the effect of rotating
   * the entire chart:
   *
   * @code
   *  chart.set_rotation(28);
   * @endcode
   *
   * The angle of rotation must be in the range `0 <= rotation <= 360`.
   *
   * This option is only available for Pie/Doughnut charts.
   */
  void set_rotation(uint16_t rotation);

  /**
   * @brief Set the Doughnut chart hole size.
   *
   * @param size The hole size as a percentage.
   *
   * The `%set_hole_size()` function is used to set the hole size of a
   * Doughnut chart:
   *
   * @code
   *  chart.set_hole_size(33);
   * @endcode
   *
   * The hole size must be in the range `10 <= size <= 90`.
   *
   * This option is only available for Doughnut charts.
   */
  void set_hole_size(uint8_t size);

  /**
   * @brief Set the font properties for a chart legend.
   *
   * @param font A pointer to a chart `chart_font_t` font struct.
   *
   * The `%legend_set_font()` function is used to set the font of a
   * chart legend:
   *
   * @code
   *  xwpp::chart_font_t font{
   *    .bold_ = true,
   *    .color_ = xwpp::color_t::BLUE,
   *  };
   *
   *  chart.legend_set_font(font);
   * @endcode
   *
   * @image html chart_legend_set_font.png
   *
   * @see @ref chart_fonts.
   */
  void legend_set_font(const std::optional<chart_font_t>& font);

  /**
   * @brief Set the gap between series in a Bar/Column chart.
   *
   * @param gap The gap between the series.  0 to 500.
   *
   * The `%set_series_gap()` function sets the gap between series in
   * Bar and Column charts.
   *
   * @code
   *  chart.set_series_gap(400);
   * @endcode
   *
   * @image html chart_gap.png
   *
   * The gap value must be in the range `0 <= gap <= 500`. The default value
   * is 150.
   *
   * This option is only available for Bar/Column charts.
   */
  void set_series_gap(uint16_t gap);

  /**
   * @brief Set the line properties for a chartarea.
   *
   * @param line A `chart_line_t` struct.
   *
   * Set the line/border properties of a chartarea. In Excel the chartarea
   * is the background area behind the chart:
   *
   * @code
   *  xwpp::chart_line_t line{.none_  = true};
   *  xwpp::chart_fill_t fill{.color_ = xwpp::color_t::RED};
   *
   *  chart.chartarea_set_line(line);
   *  chart.chartarea_set_fill(fill);
   * @endcode
   *
   * @image html chart_chartarea.png
   *
   * @see @ref chart_lines.
   */
  void chartarea_set_line(const std::optional<chart_line_t>& line);

  /**
   * @brief Set the fill properties for a chartarea.
   *
   * @param fill  A `chart_fill_t` struct.
   *
   * Set the fill properties of a chartarea:
   *
   * @code
   *  chart.chartarea_set_fill(fill);
   * @endcode
   *
   * @see The example and image above.
   *
   * @see @ref chart_fills.
   */
  void chartarea_set_fill(const std::optional<chart_fill_t>& fill);

  /**
   * @brief Set the line properties for a plotarea.
   *
   * @param line  A `chart_line_t` struct.
   *
   * Set the line/border properties of a plotarea. In Excel the plotarea is
   * the area between the axes on which the chart series are plotted:
   *
   * @code
   *  xwpp::chart_line_t line{
   *    .color_     = xwpp::color_t::RED,
   *    .width_     = 2,
   *    .dash_type_ = xwpp::chart_line_dash_type_t::DASH,
   *  };
   *  xwpp::chart_fill_t fill{
   *    .color_     = xwpp::color_t(0xFFFFC2)
   *  };
   *
   *  chart.plotarea_set_line(line);
   *  chart.plotarea_set_fill(fill);
   * @endcode
   *
   * @image html chart_plotarea.png
   *
   * @see @ref chart_lines.
   */
  void plotarea_set_line(const std::optional<chart_line_t>& line);

  /**
   * @brief Set the fill properties for a plotarea.
   *
   * @param fill A `chart_fill_t` struct.
   *
   * Set the fill properties of a plotarea:
   *
   * @code
   *  chart.plotarea_set_fill(fill);
   * @endcode
   *
   * @see The example and image above.
   *
   * @see @ref chart_fills.
   */
  void plotarea_set_fill(const std::optional<chart_fill_t>& fill);

  /**
   * @brief Set the option for displaying blank data in a chart.
   *
   * @param option The display option. A `chart_blank_t` option.
   *
   * The `%show_blanks_as()` function controls how blank data is displayed
   * in a chart:
   *
   * @code
   *  chart.show_blanks_as(xwpp::chart_blank_t::AS_CONNECTED);
   * @endcode
   *
   * The `option` parameter can have one of the following values:
   *
   * - `%chart_blank_t::AS_GAP`: Show empty chart cells as gaps in the data.
   *   This is the default option for Excel charts.
   * - `%chart_blank_t::AS_ZERO`: Show empty chart cells as zeros.
   * - `%chart_blank_t::AS_CONNECTED`: Show empty chart cells as connected.
   *   Only for charts with lines.
   */
  void show_blanks_as(chart_blank_t option);

  /**
   * @brief Display data on charts from hidden rows or columns.
   *
   * Display data that is in hidden rows or columns on the chart:
   *
   * @code
   *  chart.show_hidden_data();
   * @endcode
   */
  void show_hidden_data();

  /**
   * @brief Set a chart title formula using row and column values.
   *
   * @param sheetname The name of the worksheet that contains the cell range.
   * @param row_num   The zero indexed row number of the range.
   * @param col_num   The zero indexed column number of the range.
   *
   * The `%title_set_name_range()` function can be used to set a chart
   * title range and is an alternative to using `title_set_name()` and a
   * string formula:
   *
   * @code
   *  chart.title_set_name_range("Sheet1", 1, 0);
   * @endcode
   */
  void title_set_name_range(const std::string& sheetname, row_num_t row_num, col_num_t col_num);

  /**
   * @brief Remove one or more series from the the legend.
   *
   * @param delete_series An array of zero-indexed values to delete from series.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%legend_delete_series()` function allows you to remove/hide one
   * or more series in a chart legend (the series will still display on the chart).
   *
   * This function takes a vector of one or more zero indexed series
   * numbers.
   *
   * For example to remove the first and third zero-indexed series from the
   * legend of a chart with 3 series:
   *
   * @code
   *  std::vector<int16_t> series{0, 2};
   *
   *  chart.legend_delete_series(series);
   * @endcode
   *
   * @image html chart_legend_delete.png
   */
  void legend_delete_series(const std::vector<int16_t>& delete_series);

  /**
   * @brief Set the overlap between series in a Bar/Column chart.
   *
   * @param overlap The overlap between the series. -100 to 100.
   *
   * The `%set_series_overlap()` function sets the overlap between series
   * in Bar and Column charts.
   *
   * @code
   *  chart.set_series_overlap(chart, -50);
   * @endcode
   *
   * @image html chart_overlap.png
   *
   * The overlap value must be in the range `0 <= overlap <= 500`.
   * The default value is 0.
   *
   * This option is only available for Bar/Column charts.
   */
  void set_series_overlap(int8_t overlap);

  /**
   * @brief Get an axis pointer from a chart.
   *
   * @param axis_type The axis type (X or Y): `chart_axis_type_t`.
   *
   * The `%axis_get()` function returns a reference to a chart axis based
   * on the `chart_axis_type_t`:
   *
   * @code
   *  xwpp::chart_axis_t& x_axis = chart.axis_get(xwpp::chart_axis_type_t::TYPE_X);
   *  xwpp::chart_axis_t& y_axis = chart.axis_get(xwpp::chart_axis_type_t::TYPE_Y);
   *
   *  // Use the axis pointer in other functions.
   *  chart_axis_major_gridlines_set_visible(x_axis, true);
   *  chart_axis_major_gridlines_set_visible(y_axis, true);
   * @endcode
   */
  [[nodiscard]] chart_axis_t& axis_get(chart_axis_type_t axis_type);

  /**
   * @brief Set the manual layout of the chart plotarea.
   *
   * @param layout A pointer to a chart `chart_layout_t` struct.
   *
   * This function is used to simulate setting the manual position of the chart
   * plotarea.
   *
   * @see @ref chart_layout for more information.
   */
  void plotarea_set_layout(const std::optional<chart_layout_t>& layout);

  /**
   * @brief Set the manual layout of the chart legend.
   *
   * @param layout A pointer to a chart `chart_layout_t` struct.
   *
   * This function is used to simulate setting the manual position of the chart
   * legend.
   *
   * @see @ref chart_layout for more information.
   */
  void legend_set_layout(const std::optional<chart_layout_t>& layout);

  /**
   * @brief Set the manual position of the chart title.
   *
   * @param layout A pointer to a chart `chart_layout_t` struct.
   *
   * This function is used to simulate setting the manual position of the chart
   * title.
   *
   * @see @ref chart_layout for more information.
   */
  void title_set_layout(const std::optional<chart_layout_t>& layout);

  /**
   * @brief Allow the chart title to overlay the chart.
   *
   * @param overlay Turn off/on the overlay. (0/1)
   *
   * This option allows the chart title to overlay the chart when the
   * `title_set_layout()` function.
   */
  void title_set_overlay(bool overlay);

  /**
   * @brief Set the pattern properties for a plotarea.
   *
   * @param pattern A `chart_pattern_t` struct.
   *
   * Set the pattern properties of a plotarea:
   *
   * @code
   *  chart.plotarea_set_pattern(series1, pattern);
   * @endcode
   *
   * @see `chart_pattern_type_t` and @ref chart_patterns.
   */
  void plotarea_set_pattern(const std::optional<chart_pattern_t>& pattern);

  // TODO Add documentation
  void set_table_font(const std::optional<chart_font_t>& font);

  /**
   * @brief Turn off an automatic chart title.
   *
   * In general in Excel a chart title isn't displayed unless the user
   * explicitly adds one. However, Excel adds an automatic chart title to charts
   * with a single series and a user defined series name. The
   * `title_off()` function allows you to turn off this automatic chart
   * title:
   *
   * @code
   *  chart.title_off();
   * @endcode
   */
  void title_off();

  /**
   * @brief Set the pattern properties for a chartarea.
   *
   * @param pattern A `chart_pattern_t` struct.
   *
   * Set the pattern properties of a chartarea:
   *
   * @code
   *  chart.chartarea_set_pattern(pattern);
   * @endcode
   *
   * @see `chart_pattern_type_t` and @ref chart_patterns.
   */
  void chartarea_set_pattern(const std::optional<chart_pattern_t>& pattern);

  // TODO Only used by packager, should be private.
  [[nodiscard]] std::string assemble_xml_file();

  // TODO For test
  void set_axis_ids(uint32_t axis_id_1, uint32_t axis_id_2);

  // TODO Set to public as chart_axis_set_name access to it.
  // To be set private again, and chart_axis_set_name reworked
  // NOLINTNEXTLINE(cppcoreguidelines-non-private-member-variables-in-classes,misc-non-private-member-variables-in-classes)
  chart_axis_t x_axis_{.default_num_format_ = "General",
                       .major_gridlines_    = {.visible_ = false},
                       .axis_position_      = chart_position_t::BOTTOM};
  // NOLINTNEXTLINE(cppcoreguidelines-non-private-member-variables-in-classes,misc-non-private-member-variables-in-classes)
  chart_axis_t y_axis_{
    .default_num_format_ = "General", .major_gridlines_ = {.visible_ = true}, .axis_position_ = chart_position_t::LEFT};

private:
  friend class worksheet_t;  // TODO
  friend class chartsheet_t; // TODO
  friend class workbook_t;

  // Initialisation functions
  void initialize_column_chart(chart_type_t type);
  void initialize_bar_chart(chart_type_t type);
  void initialize_area_chart(chart_type_t type);
  void initialize_line_chart(chart_type_t type);
  void initialize_doughnut_chart();
  void initialize_pie_chart();
  void initialize_radar_chart(chart_type_t type);
  void initialize_scatter_chart(chart_type_t type);
  void initialize(chart_type_t type);

  // Functions to generate XML
  [[nodiscard]] static std::string write_bar_chart(chart_t& chart);
  [[nodiscard]] static std::string write_column_chart(chart_t& chart);
  [[nodiscard]] static std::string write_plot_area(chart_t& chart);
  [[nodiscard]] static std::string write_area_chart(chart_t& chart);
  [[nodiscard]] static std::string write_line_chart(chart_t& chart);
  [[nodiscard]] static std::string write_doughnut_chart(chart_t& chart);
  [[nodiscard]] static std::string write_pie_plot_area(chart_t& chart);
  [[nodiscard]] static std::string write_pie_chart(chart_t& chart);
  [[nodiscard]] static std::string write_radar_chart(chart_t& chart);
  [[nodiscard]] static std::string write_scatter_plot_area(chart_t& chart);
  [[nodiscard]] static std::string write_scatter_chart(chart_t& chart);
  [[nodiscard]] static std::string write_chart_space();
  [[nodiscard]] static std::string write_lang();
  [[nodiscard]] std::string write_style() const;
  [[nodiscard]] static std::string write_protection();
  [[nodiscard]] static std::string write_print_settings();
  [[nodiscard]] static std::string write_header_footer();
  [[nodiscard]] static std::string write_page_margins();
  [[nodiscard]] static std::string write_page_setup();
  [[nodiscard]] std::string write_chart();
  [[nodiscard]] std::string write_chart_title() const;
  [[nodiscard]] static std::string write_auto_title_deleted();
  [[nodiscard]] static std::string write_tx_pr_pie(bool is_horizontal, const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_p_pie(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_p_pr_pie(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_title(const chart_title_t& title);
  [[nodiscard]] static std::string write_title_rich(const chart_title_t& title);
  [[nodiscard]] static std::string write_tx_rich(const std::string& name, bool is_horizontal,
                                                 const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_rich(const std::string& name, const std::optional<chart_font_t>& font,
                                              bool is_horizontal, bool ignore_rich_pr);
  [[nodiscard]] static std::string write_a_body_pr(int32_t rotation, bool is_horizontal);
  [[nodiscard]] static std::string write_a_lst_style();
  [[nodiscard]] static std::string write_a_p_rich(const std::string& name, const std::optional<chart_font_t>& font,
                                                  bool ignore_rich_pr);
  [[nodiscard]] static std::string write_a_p_pr_rich(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_def_rpr(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_r(const std::string& name, const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_r_pr(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_t(const std::string& name);
  [[nodiscard]] static std::string write_layout(const std::optional<chart_layout_t>& layout);
  [[nodiscard]] static std::string write_manual_layout(const chart_layout_t& layout);
  [[nodiscard]] static std::string write_overlay();
  [[nodiscard]] static std::string write_bar_dir(const std::string& type);
  [[nodiscard]] static std::string write_grouping(chart_grouping_t grouping);
  [[nodiscard]] static std::string write_gap_width(uint16_t gap);
  [[nodiscard]] static std::string write_overlap(int8_t overlap);
  [[nodiscard]] static std::string write_axis_id(uint32_t axis_id);
  [[nodiscard]] static std::string write_axis_ids(chart_t& chart);
  [[nodiscard]] static std::string write_ser(chart_t& chart, chart_series_t& series);
  [[nodiscard]] static std::string write_idx(uint16_t index);
  [[nodiscard]] static std::string write_order(uint16_t index);
  [[nodiscard]] static std::string write_series_name(const chart_series_t& series);
  [[nodiscard]] static std::string write_tx_value(const std::string& name);
  [[nodiscard]] static std::string write_v_str(const std::string& str);
  [[nodiscard]] static std::string write_tx_formula(const chart_title_t& title);
  [[nodiscard]] static std::string write_str_ref(const series_range_t& range);
  [[nodiscard]] static std::string write_f(const std::string& formula);
  [[nodiscard]] static std::string write_sp_pr(const std::optional<chart_line_t>& line,
                                               const std::optional<chart_fill_t>& fill,
                                               const std::optional<chart_pattern_t>& pattern);
  [[nodiscard]] static std::string write_a_no_fill();
  [[nodiscard]] static std::string write_a_solid_fill(color_t color, uint8_t transparency);
  [[nodiscard]] static std::string write_a_srgb_clr(color_t color, uint8_t transparency);
  [[nodiscard]] static std::string write_a_alpha(uint8_t transparency);
  [[nodiscard]] static std::string write_a_patt_fill(const chart_pattern_t& pattern);
  [[nodiscard]] static std::string write_a_fg_clr(color_t color);
  [[nodiscard]] static std::string write_a_bg_clr(color_t color);
  [[nodiscard]] static std::string write_a_ln(const chart_line_t& line);
  [[nodiscard]] static std::string write_a_prst_dash(chart_line_dash_type_t dash_type);
  [[nodiscard]] static std::string write_str_cache(const series_range_t& range);
  [[nodiscard]] static std::string write_pt_count(uint16_t num_data_points);
  [[nodiscard]] static std::string write_pt(uint16_t index, const series_data_point_t& data_point);
  [[nodiscard]] static std::string write_v_num(double number);
  [[nodiscard]] static std::string write_cat(chart_t& chart, const chart_series_t& series);
  [[nodiscard]] static std::string write_data_cache(const series_range_t& range, bool has_string_cache);
  [[nodiscard]] static std::string write_num_ref(const series_range_t& range);
  [[nodiscard]] static std::string write_num_cache(const series_range_t& range);
  [[nodiscard]] static std::string write_format_code();
  [[nodiscard]] static std::string write_num_pt(uint16_t index, const series_data_point_t& data_point);
  [[nodiscard]] static std::string write_val_axis(chart_t& chart);
  [[nodiscard]] static std::string write_scaling(bool reverse, bool has_min, double min, bool has_max, double max,
                                                 uint16_t log_base);
  [[nodiscard]] static std::string write_max(double max);
  [[nodiscard]] static std::string write_min(double min);
  [[nodiscard]] static std::string write_log_base(uint16_t log_base);
  [[nodiscard]] static std::string write_orientation(bool reverse);
  [[nodiscard]] static std::string write_val(const chart_series_t& series);
  [[nodiscard]] static std::string write_marker(const chart_t& chart, std::optional<chart_marker_t>& marker);
  [[nodiscard]] static std::string write_cat_axis(chart_t& chart);
  [[nodiscard]] static std::string write_delete();
  [[nodiscard]] static std::string write_axis_pos(chart_position_t position, bool reverse);
  [[nodiscard]] static std::string write_title_formula(const chart_title_t& title);
  [[nodiscard]] static std::string write_tx_pr(bool is_horizontal, const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_p_formula(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_a_p_pr_formula(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_axis_font(const std::optional<chart_font_t>& font);
  [[nodiscard]] static std::string write_major_gridlines(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_minor_gridlines(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_cat_number_format(const chart_t& chart, const chart_axis_t& axis);
  [[nodiscard]] static std::string write_tick_label_pos(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_cross_axis(uint32_t axis_id);
  [[nodiscard]] static std::string write_crosses(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_crosses_at(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_auto();
  [[nodiscard]] static std::string write_label_align(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_label_offset();
  [[nodiscard]] static std::string write_number_format(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_cross_between(const chart_t& chart, chart_axis_tick_position_t position);
  [[nodiscard]] static std::string write_legend_pos(const std::string& position);
  [[nodiscard]] std::string write_legend_entry(size_t index);
  [[nodiscard]] std::string write_legend();
  [[nodiscard]] std::string write_plot_vis_only() const;
  [[nodiscard]] static std::string write_drop_lines(const chart_t& chart);
  [[nodiscard]] static std::string write_d_lbls(const chart_series_t& series);
  [[nodiscard]] static std::string write_custom_label_str(const chart_series_t& series,
                                                          const chart_custom_label_t& data_label);
  [[nodiscard]] static std::string write_custom_labels(const chart_series_t& series);
  [[nodiscard]] static std::string write_d_lbl_pos(chart_label_position_t position);
  [[nodiscard]] static std::string write_show_val();
  [[nodiscard]] static std::string write_show_cat_name();
  [[nodiscard]] static std::string write_show_ser_name();
  [[nodiscard]] static std::string write_custom_label_formula(const chart_series_t& series,
                                                              const chart_custom_label_t& data_label);
  [[nodiscard]] static std::string write_custom_label_format_only(const chart_custom_label_t& data_label);
  [[nodiscard]] static std::string write_label_num_fmt(const std::string& format);
  [[nodiscard]] static std::string write_show_legend_key();
  [[nodiscard]] static std::string write_separator(chart_label_separator_t separator);
  [[nodiscard]] static std::string write_show_percent();
  [[nodiscard]] static std::string write_show_leader_lines();
  [[nodiscard]] static std::string write_a_end_para_rpr();
  [[nodiscard]] static std::string write_d_table(const chart_t& chart);
  [[nodiscard]] static std::string write_show_horz_border(bool value);
  [[nodiscard]] static std::string write_show_vert_border(bool value);
  [[nodiscard]] static std::string write_show_outline(bool value);
  [[nodiscard]] static std::string write_show_keys(bool value);
  [[nodiscard]] static std::string write_hi_low_lines(const chart_t& chart);
  [[nodiscard]] static std::string write_up_down_bars(const chart_t& chart);
  [[nodiscard]] static std::string write_marker_value();
  [[nodiscard]] static std::string write_up_bars(const std::optional<chart_line_t>& line,
                                                 const std::optional<chart_fill_t>& fill);
  [[nodiscard]] static std::string write_down_bars(const std::optional<chart_line_t>& line,
                                                   const std::optional<chart_fill_t>& fill);
  [[nodiscard]] static std::string write_symbol(chart_marker_type_t type);
  [[nodiscard]] static std::string write_marker_size(uint8_t size);
  [[nodiscard]] static std::string write_error_bars(const chart_series_t& series);
  [[nodiscard]] static std::string write_err_bars(const series_error_bars_t& error_bars);
  [[nodiscard]] static std::string write_err_dir(bool is_x);
  [[nodiscard]] static std::string write_err_bar_type(chart_error_bar_direction_t direction);
  [[nodiscard]] static std::string write_err_val_type(chart_error_bar_type_t type);
  [[nodiscard]] static std::string write_no_end_cap();
  [[nodiscard]] static std::string write_error_val(double value);
  [[nodiscard]] static std::string write_trendline(const chart_series_t& series);
  [[nodiscard]] static std::string write_name(const std::string& name);
  [[nodiscard]] static std::string write_trendline_type(chart_trendline_type_t type);
  [[nodiscard]] static std::string write_period(uint8_t value);
  [[nodiscard]] static std::string write_forward(double value);
  [[nodiscard]] static std::string write_backward(double value);
  [[nodiscard]] static std::string write_intercept(double value);
  [[nodiscard]] static std::string write_disp_rsqr();
  [[nodiscard]] static std::string write_disp_eq();
  [[nodiscard]] static std::string write_trendline_lbl();
  [[nodiscard]] static std::string write_invert_if_negative(const chart_series_t& series);
  [[nodiscard]] static std::string write_points(const chart_t& chart, const chart_series_t& series);
  [[nodiscard]] static std::string write_d_pt(const chart_t& chart, const chart_point_t& point, uint16_t index);
  [[nodiscard]] static std::string write_vary_colors();
  [[nodiscard]] static std::string write_first_slice_ang(const chart_t& chart);
  [[nodiscard]] static std::string write_hole_size(const chart_t& chart);
  [[nodiscard]] static std::string write_radar_style(const chart_t& chart);
  [[nodiscard]] static std::string write_major_tick_mark(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_minor_tick_mark(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_cat_val_axis(chart_t& chart);
  [[nodiscard]] static std::string write_xval_ser(chart_t& chart, chart_series_t& series);
  [[nodiscard]] static std::string write_x_val(const chart_series_t& series);
  [[nodiscard]] static std::string write_y_val(const chart_series_t& series);
  [[nodiscard]] static std::string write_smooth(bool smooth);
  [[nodiscard]] static std::string write_major_unit(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_minor_unit(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_disp_units(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_scatter_style(const chart_t& chart);
  [[nodiscard]] static std::string write_tick_label_skip(const chart_axis_t& axis);
  [[nodiscard]] static std::string write_tick_mark_skip(const chart_axis_t& axis);
  [[nodiscard]] std::string write_disp_blanks_as() const;
  [[nodiscard]] static std::string write_layout_target();
  [[nodiscard]] static std::string write_layout_mode(const std::string& mode);
  [[nodiscard]] static std::string write_layout_dimension(const std::string& dimension, double value);

  static void add_axis_ids(chart_t& chart);
  static void adjust_max_crossing(chart_t& chart);

  chart_type_t type_;
  chart_subtype_t subtype_{chart_subtype_t::NONE};
  uint16_t series_index_{0};

  std::function<std::string(chart_t& chart)> write_chart_type_;
  std::function<std::string(chart_t& chart)> write_plot_area_;

  chart_title_t title_;
  uint32_t id_{0};
  uint32_t axis_id_1_{0};
  uint32_t axis_id_2_{0};
  // TODO ?   uint32_t axis_id_3;
  // TODO ?   uint32_t axis_id_4;
  bool in_use_{false};
  chart_type_t chart_group_{chart_type_t::NONE};
  bool cat_has_num_fmt_{false};
  bool is_chartsheet_{false};
  bool has_horiz_cat_axis_{false};
  bool has_horiz_val_axis_{true};
  uint8_t style_id_{2}; // TODO Constant for default style
  uint16_t rotation_{0};
  uint16_t hole_size_{50};
  // TODO ?   uint8_t no_title;
  bool has_overlap_{false};
  int8_t overlap_y1_{0};
  // TODO ?   int8_t overlap_y2;
  uint16_t gap_y1_{DEFAULT_GAP};
  uint16_t gap_y2_{DEFAULT_GAP};
  chart_grouping_t grouping_{chart_grouping_t::CLUSTERED};
  chart_axis_tick_position_t default_cross_between_{chart_axis_tick_position_t::DEFAULT};
  chart_legend_t legend_{.position_ = chart_legend_position_t::RIGHT};
  std::vector<int16_t> delete_series_;
  std::optional<chart_marker_t> default_marker_;
  std::optional<chart_line_t> chartarea_line_;
  std::optional<chart_fill_t> chartarea_fill_;
  std::optional<chart_pattern_t> chartarea_pattern_;
  std::optional<chart_line_t> plotarea_line_;
  std::optional<chart_fill_t> plotarea_fill_;
  std::optional<chart_layout_t> plotarea_layout_;
  std::optional<chart_pattern_t> plotarea_pattern_;
  bool has_drop_lines_{false};
  std::optional<chart_line_t> drop_lines_line_;
  bool has_high_low_lines_{false};
  std::optional<chart_line_t> high_low_lines_line_;
  std::list<chart_series_t> series_list_;
  bool has_table_{false};
  bool has_table_vertical_{false};
  bool has_table_horizontal_{false};
  bool has_table_outline_{false};
  bool has_table_legend_keys_{false};
  std::optional<chart_font_t> table_font_;
  chart_blank_t show_blanks_as_{chart_blank_t::AS_GAP};
  bool show_hidden_data_{false};
  bool has_up_down_bars_{false};
  std::optional<chart_line_t> up_bar_line_;
  std::optional<chart_line_t> down_bar_line_;
  std::optional<chart_fill_t> up_bar_fill_;
  std::optional<chart_fill_t> down_bar_fill_;
  chart_label_position_t default_label_position_{chart_label_position_t::DEFAULT};
  bool is_protected_{false};

  static const uint16_t DEFAULT_GAP{501};
};

/**
 * @brief Set a series "categories" range using row and column values.
 *
 * @param series    A series object created via `chart_t::add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range.
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `chart_t::add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%chart_series_set_categories()` function is an alternative method that
 * is easier to generate programmatically. It requires that you set the
 * `categories` and `values` parameters in `chart_add_series()`to ann empty
 * string and then set them using row and column values in
 * `chart_series_set_categories()` and `chart_series_set_values()`:
 *
 * @code
 *  xwpp::chart_series_t& series = chart.add_series("", "");
 *
 *  // Configure the series ranges programmatically.
 *  chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0);  // "=Sheet1!$A$2:$A$7"
 *  chart_series_set_values(series, "Sheet1", 1, 2, 6, 2);      // "=Sheet1!$C$2:$C$7"
 * @endcode
 */
void chart_series_set_categories(chart_series_t& series, const std::string& sheetname, row_num_t first_row,
                                 col_num_t first_col, row_num_t last_row, col_num_t last_col);

/**
 * @brief Set a series "values" range using row and column values.
 *
 * @param series    A series object created via `chart_t::add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range.
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `chart_t::add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%chart_series_set_values()` function is an alternative method that is
 * easier to generate programmatically.
 *
 * @see the documentation for `chart_series_set_categories()` above.
 */
void chart_series_set_values(chart_series_t& series, const std::string& sheetname, row_num_t first_row,
                             col_num_t first_col, row_num_t last_row, col_num_t last_col);

// TODO Add doc.
void set_range(series_range_t& range, const std::string& sheetname, row_num_t first_row, col_num_t first_col,
               row_num_t last_row, col_num_t last_col);

// TODO To move out chart_t
/**
 * @brief Set a series name formula using row and column values.
 *
 * @param series    A series object created via `chart_t::add_series()`.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row_num   The zero indexed row number of the range.
 * @param col_num   The zero indexed column number of the range.
 *
 * The `%chart_series_set_name_range()` function can be used to set a series
 * name range and is an alternative to using `chart_series_set_name()` and a
 * string formula:
 *
 * @code
 *  xwpp::chart_series_t& series = chart.add_series("", "=Sheet1!$B$2:$B$7");
 *
 *  chart_series_set_name_range(series, "Sheet1", 0, 2); // "=Sheet1!$C$1"
 * @endcode
 */
void chart_series_set_name_range(chart_series_t& series, const std::string& sheetname, row_num_t row_num,
                                 col_num_t col_num);

/**
 * @brief Set the line properties for a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param line   A `chart_line_t` struct.
 *
 * Set the line/border properties of a chart series:
 *
 * @code
 *  xwpp::chart_line_t line{.color_ = xwpp::color_t::RED};
 *
 *  chart_series_set_line(series1, line);
 *  chart_series_set_line(series2, line);
 *  chart_series_set_line(series3, line);
 * @endcode
 *
 * @image html chart_series_set_line.png
 *
 * @see @ref chart_lines.
 */
void chart_series_set_line(chart_series_t& series, const std::optional<chart_line_t>& line);

/**
 * @brief Set the fill properties for a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param fill   A `chart_fill_t` struct.
 *
 * Set the fill properties of a chart series:
 *
 * @code
 *  xwpp::chart_fill_t fill1{.color_ = xwpp::color_t::RED};
 *  xwpp::chart_fill_t fill2{.color_ = xwpp::color_t::YELLOW};
 *  xwpp::chart_fill_t fill3{.color_ = xwpp::color_t::GREEN};
 *
 *  chart_series_set_fill(series1, fill1);
 *  chart_series_set_fill(series2, fill2);
 *  chart_series_set_fill(series3, fill3);
 * @endcode
 *
 * @image html chart_series_set_fill.png
 *
 * @see @ref chart_fills.
 */
void chart_series_set_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill);

/**
 * @brief Invert the fill color for negative series values.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * Invert the fill color for negative values. Usually only applicable to
 * column and bar charts.
 *
 * @code
 *  chart_series_set_invert_if_negative(series);
 * @endcode
 */
void chart_series_set_invert_if_negative(chart_series_t& series);

/**
 * @brief Set the pattern properties for a chart series.
 *
 * @param series  A series object created via `chart_t::add_series()`.
 * @param pattern A `chart_pattern_t` struct.
 *
 * Set the pattern properties of a chart series:
 *
 * @code
 *  xwpp::chart_pattern_t pattern1{
 *    .type_     = xwpp::chart_pattern_type_t::SHINGLE,
 *    .fg_color_ = xwpp::color_t(0x804000),
 *    .bg_color_ = xwpp::color_t(0xC68C53),
 *  };
 *
 *  xwpp::chart_pattern_t pattern2{
 *    .type_     = xwpp::chart_pattern_type_t::HORIZONTAL_BRICK,
 *    .fg_color_ = xwpp::color_t(0xB30000),
 *    .bg_color_ = xwpp::color_t(0xFF6666),
 *  };
 *
 *  chart_series_set_pattern(series1, pattern1);
 *  chart_series_set_pattern(series2, pattern2);
 * @endcode
 *
 * @image html chart_pattern.png
 *
 * @see `chart_pattern_type_t` and @ref chart_patterns.
 */
void chart_series_set_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern);

/**
 * @brief Set the size of a data marker for a series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param size   The size of the marker.
 *
 * The `%chart_series_set_marker_size()` function is used to specify the
 * size of the series marker:
 *
 * @code
 *  chart.series_set_marker_type(series, xwpp::chart_marker_type_t::CIRCLE);
 *  chart.series_set_marker_size(series, 10);
 * @endcode
 *
 * @image html chart_series_set_marker_size.png
 */
void chart_series_set_marker_size(chart_series_t& series, uint8_t size);

/**
 * @brief Set the line properties for a chart series marker.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param line   A `chart_line_t` struct.
 *
 * Set the line/border properties of a chart marker:
 *
 * @code
 *  xwpp::chart_line_t line{.color_ = xwpp::color_t::BLACK};
 *  xwpp::chart_fill_t fill{.color_ = xwpp::color_t::RED};
 *
 *  chart.series_set_marker_type(series, xwpp::chart_marker_type_t::SQUARE);
 *  chart.series_set_marker_size(series, 8);
 *
 *  chart_series_set_marker_line(series, line);
 *  chart_series_set_marker_fill(series, fill);
 * @endcode
 *
 * @image html chart_marker2.png
 *
 * @see @ref chart_lines.
 */
void chart_series_set_marker_line(chart_series_t& series, const std::optional<chart_line_t>& line);

/**
 * @brief Set the fill properties for a chart series marker.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param fill   A `chart_fill_t` struct.
 *
 * Set the fill properties of a chart marker:
 *
 * @code
 *  chart_series_set_marker_fill(series, fill);
 * @endcode
 *
 * @see @ref chart_fills.
 */
void chart_series_set_marker_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill);

/**
 * @brief Set the pattern properties for a chart series marker.
 *
 * @param series  A series object created via `chart_t::add_series()`.
 * @param pattern A `chart_pattern_t` struct.
 *
 * Set the pattern properties of a chart marker:
 *
 * @code
 *  chart_series_set_marker_pattern(series, pattern);
 * @endcode
 *
 * @see `chart_pattern_type_t` and @ref chart_patterns.
 */
void chart_series_set_marker_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern);

/**
 * @brief Set the formatting for points in the series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param points A vector of `chart_point_t` pointers.
 *
 * @throw xwpp::xwpp_exception_t.
 *
 * In general formatting is applied to an entire series in a chart. However,
 * it is occasionally required to format individual points in a series. In
 * particular this is required for Pie/Doughnut charts where each segment is
 * represented by a point.
 *
 * @dontinclude chart_pie_colors.cpp
 * @skip Add the data series
 * @until series_set_points
 *
 * @image html chart_points1.png
 *
 * @see @ref chart_points
 */
void series_set_points(chart_series_t& series, const std::vector<chart_point_t>& points);

/**
 * @brief Smooth a line or scatter chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param smooth Turn off/on the line smoothing. (0/1)
 *
 * The `chart_series_set_smooth()` function is used to set the smooth property
 * of a line series. It is only applicable to the line and scatter chart
 * types:
 *
 * @code
 *  chart_series_set_smooth(series2, true);
 * @endcode
 *
 * @image html chart_smooth.png
 */
void chart_series_set_smooth(chart_series_t& series, bool smooth);

// TODO Member function of series
/**
 * @brief Add data labels to a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_labels()` function is used to turn on data labels
 * for a chart series. Data labels indicate the values of the plotted data
 * points.
 *
 * @code
 *  chart_series_set_labels(series);
 * @endcode
 *
 * @image html chart_data_labels1.png
 *
 * By default data labels are displayed in Excel with only the values shown:
 *
 * @image html chart_data_labels2.png
 *
 * However, it is possible to configure other display options, as shown
 * in the functions below.
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels(chart_series_t& series);

/**
 * @brief Set the display options for the labels of a data series.
 *
 * @param series        A series object created via `chart_t::add_series()`.
 * @param show_name     Turn on/off the series name in the label caption.
 * @param show_category Turn on/off the category name in the label caption.
 * @param show_value    Turn on/off the value in the label caption.
 *
 * The `%chart_series_set_labels_options()` function is used to set the
 * parameters that are displayed in the series data label:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_options(series, true, true, true);
 * @endcode
 *
 * @image html chart_data_labels3.png
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_options(chart_series_t& series, bool show_name, bool show_category, bool show_value);

/** @brief Set the properties for data labels in a series.
 *
 * @param series      A series object created via `chart_t::add_series()`.
 * @param data_labels A vector of `chart_data_label_t` pointers.
 *
 * @throw xwpp::xwpp_exception_t.
 *
 * The `%chart_series_set_labels_custom()` function is used to set the properties
 * for data labels in a series. It can also be used to delete individual data
 * labels in a series.
 *
 * In general properties are set for all the data labels in a chart
 * series. However, it is also possible to set properties for individual data
 * labels in a series using `%chart_series_set_labels_custom()`.
 *
 * @code
 *  // Add the series data labels.
 *  chart_series_set_labels(series);
 *
 *  // Create some custom labels.
 *  std::vector<xwpp::chart_data_label_t> data_labels{
 *    {.value = "Jan"},
 *    {.value = "Feb"},
 *    {.value = "Mar"},
 *    {.value = "Apr"},
 *    {.value = "May"},
 *    {.value = "Jun"},
 *  };
 *
 *  // Set the custom labels.
 *  chart_series_set_labels_custom(series, data_labels);
 * @endcode
 *
 * @image html chart_data_labels18.png
 *
 * @see @ref chart_custom_labels.
 */
void chart_series_set_labels_custom(chart_series_t& series, const std::vector<chart_data_label_t>& data_labels);

/**
 * @brief Set the separator for the data label captions.
 *
 * @param series    A series object created via `chart_t::add_series()`.
 * @param separator The separator for the data label options: `chart_label_separator_t`.
 *
 * The `%chart_series_set_labels_separator()` function is used to change the
 * separator between multiple data label items. The default options is a comma
 * separator as shown in the previous example.
 *
 * The available options are:
 *
 * - `%chart_label_separator_t::SEMICOLON`: semicolon separator.
 * - `%chart_label_separator_t::PERIOD`: a period (dot) separator.
 * - `%chart_label_separator_t::NEWLINE`: a newline separator.
 * - `%chart_label_separator_t::SPACE`: a space separator.
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_options(series, true, true, true);
 *  chart_series_set_labels_separator(series, xwpp::chart_label_separator_t::NEWLINE);
 * @endcode
 *
 * @image html chart_data_labels4.png
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_separator(chart_series_t& series, chart_label_separator_t separator);

/**
 * @brief Set the data label position for a series.
 *
 * @param series   A series object created via `chart_t::add_series()`.
 * @param position The data label position: `chart_label_position_t`.
 *
 * The `%chart_series_set_labels_position()` function sets the position of
 * the labels in the data series:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_position(series, xwpp::chart_label_position_t::ABOVE);
 * @endcode
 *
 * @image html chart_data_labels5.png
 *
 * In Excel the allowable data label positions vary for different chart
 * types. The allowable, and default, positions are:
 *
 * | Position                                | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
 * | :-------------------------------------- | :------------ | :------------ | :------------ | :------------ |
 * | `%chart_label_position_t::CENTER`       | Yes           | Yes           | Yes           | Yes (default) |
 * | `%chart_label_position_t::RIGHT`        | Yes (default) |               |               |               |
 * | `%chart_label_position_t::LEFT`         | Yes           |               |               |               |
 * | `%chart_label_position_t::ABOVE`        | Yes           |               |               |               |
 * | `%chart_label_position_t::BELOW`        | Yes           |               |               |               |
 * | `%chart_label_position_t::INSIDE_BASE`  |               | Yes           |               |               |
 * | `%chart_label_position_t::INSIDE_END`   |               | Yes           | Yes           |               |
 * | `%chart_label_position_t::OUTSIDE_END`  |               | Yes (default) | Yes           |               |
 * | `%chart_label_position_t::BEST_FIT`     |               |               | Yes (default) |               |
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_position(chart_series_t& series, chart_label_position_t position);

/**
 * @brief Set leader lines for Pie and Doughnut charts.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_labels_leader_line()` function  is used to turn on
 * leader lines for the data label of a series. It is mainly used for pie
 * or doughnut charts:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_leader_line(series);
 * @endcode
 *
 * @note Even when leader lines are turned on they aren't automatically
 * visible in Excel. Due to an Excel limitation (or design) leader
 * lines only appear if the data label is moved manually or if
 * the data labels are very close and need to be adjusted automatically.
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_leader_line(chart_series_t& series);

/**
 * @brief Set the legend key for a data label in a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_labels_legend()` function is used to set the
 * legend key for a data series:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_legend(series);
 * @endcode
 *
 * @image html chart_data_labels6.png
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_legend(chart_series_t& series);

/**
 * @brief Set the percentage for a Pie/Doughnut data point.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_labels_percentage()` function is used to turn on
 * the display of data labels as a percentage for a series. It is mainly
 * used for pie charts:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_options(series, false, false, false);
 *  chart_series_set_labels_percentage(series);
 * @endcode
 *
 * @image html chart_data_labels7.png
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_percentage(chart_series_t& series);

/**
 * @brief Set the number format for chart data labels in a series.
 *
 * @param series     A series object created via `chart_t::add_series()`.
 * @param num_format The number format string.
 *
 * The `%chart_series_set_labels_num_format()` function is used to set the
 * number format for data labels:
 *
 * @code
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_num_format(series, "$0.00");
 * @endcode
 *
 * @image html chart_data_labels8.png
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `format_t::set_num_format()`.
 *
 * @see @ref chart_labels.
 */
void chart_series_set_labels_num_format(chart_series_t& series, std::string_view num_format);

/**
 * @brief Set the font properties for chart data labels in a series
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param font   A pointer to a chart `chart_font_t` font struct.
 *
 * The `%chart_series_set_labels_font()` function is used to set the font
 * for data labels:
 *
 * @code
 *  xwpp::chart_font_t font{
 *    .name_ = "Consolas",
 *    .color_ = xwpp::color_t::RED,
 *  };
 *
 *  chart_series_set_labels(series);
 *  chart_series_set_labels_font(series, font);
 * @endcode
 *
 * @image html chart_data_labels9.png
 *
 * @see @ref chart_fonts and @ref chart_labels.
 */
void chart_series_set_labels_font(chart_series_t& series, const std::optional<chart_font_t>& font);

/**
 * @brief Set the line properties for the data labels in a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param line   A `chart_line_t` struct.
 *
 * Set the line/border properties of the data labels in a chart series:
 *
 * @code
 *  xwpp::chart_line_t line{.color_ = xwpp::color_t::RED};
 *  xwpp::chart_fill_t fill{.color_ = xwpp::color_t::YELLOW};
 *
 *  chart_series_set_labels_line(series, line);
 *  chart_series_set_labels_fill(series, fill);
 * @endcode
 *
 * @image html chart_data_labels24.png
 *
 * @see @ref chart_lines and @ref chart_labels.
 */
void chart_series_set_labels_line(chart_series_t& series, const std::optional<chart_line_t>& line);

/**
 * @brief Set the fill properties for the data labels in a chart series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param fill   A `chart_fill_t` struct.
 *
 * Set the fill properties of the data labels in a chart series:
 *
 * @code
 *  xwpp::chart_fill_t fill{.color_ = xwpp::color_t::YELLOW};
 *
 *  chart_series_set_labels_fill(series, fill);
 * @endcode
 *
 * @see @ref chart_fills and @ref chart_labels.
 */
void chart_series_set_labels_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill);

/**
 * @brief Set the pattern properties for the data labels in a chart series.
 *
 * @param series  A series object created via `chart_t::add_series()`.
 * @param pattern A `chart_pattern_t` struct.
 *
 * Set the pattern properties of the data labels in a chart series:
 *
 * @code
 *  chart_series_set_labels_pattern(series, pattern);
 * @endcode
 *
 * @see `chart_pattern_type_t` and @ref chart_patterns.
 */
void chart_series_set_labels_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern);

/**
 * @brief Turn on a trendline for a chart data series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param type   The type of trendline: `chart_trendline_type_t`.
 * @param value  The order/period value for polynomial and moving average trendlines.
 *
 * A trendline can be added to a chart series to indicate trends in the data
 * such as a moving average or a polynomial fit. The trendlines types are
 * shown in the following Excel dialog:
 *
 * @image html chart_trendline0.png
 *
 * The `%series_set_trendline()` function turns on these trendlines for
 * a data series:
 *
 * @code
 *  chart = workbook.add_chart(xwpp::chart_type_t::LINE);
 *  series = chart.add_series("", "Sheet1!$A$1:$A$6");
 *
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 * @endcode
 *
 * @image html chart_trendline2.png
 *
 * The `value` parameter corresponds to *order* for a polynomial trendline
 * and *period* for a Moving Average trendline. It both cases it must be >= 2.
 * The `value` parameter  is ignored for all other trendlines:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::AVERAGE, 2);
 * @endcode
 *
 * @image html chart_trendline3.png
 *
 * The allowable values for the the trendline `type` are:
 *
 * - `chart_trendline_type_t::LINEAR`: Linear trendline.
 * - `chart_trendline_type_t::LOG`: Logarithm trendline.
 * - `chart_trendline_type_t::POLY`: Polynomial trendline. The `value`
 *   parameter corresponds to *order*.
 * - `chart_trendline_type_t::POWER`: Power trendline.
 * - `chart_trendline_type_t::EXP`: Exponential trendline.
 * - `chart_trendline_type_t::AVERAGE`: Moving Average trendline. The `value`
 *   parameter corresponds to *period*.
 *
 * Other trendline options, such as those shown in the following Excel
 * dialog, can be set using the functions below.
 *
 * @image html chart_trendline1.png
 *
 * @see @ref chart_trendlines.
 */
void series_set_trendline(chart_series_t& series, chart_trendline_type_t type, uint8_t value);

/**
 * @brief Set the trendline forecast for a chart data series.
 *
 * @param series   A series object created via `chart_t::add_series()`.
 * @param forward  The forward period.
 * @param backward The backwards period.
 *
 * The `%series_set_trendline_forecast()` function sets the forward
 * and backward forecast periods for the trendline:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_forecast(series, 0.5, 0.5);
 * @endcode
 *
 * @image html chart_trendline4.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * @see @ref chart_trendlines.
 */
void chart_series_set_trendline_forecast(chart_series_t& series, double forward, double backward);

/**
 * @brief Display the equation of a trendline for a chart data series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_trendline_equation()` function displays the
 * equation of the trendline on the chart:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_equation(series);
 * @endcode
 *
 * @image html chart_trendline5.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * @see @ref chart_trendlines.
 */
void chart_series_set_trendline_equation(chart_series_t& series);

/**
 * @brief Display the R squared value of a trendline for a chart data series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 *
 * The `%chart_series_set_trendline_r_squared()` function displays the
 * R-squared value for the trendline on the chart:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_r_squared(series);
 * @endcode
 *
 * @image html chart_trendline6.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * @see @ref chart_trendlines.
 */
void chart_series_set_trendline_r_squared(chart_series_t& series);

/**
 * @brief Set the trendline Y-axis intercept for a chart data series.
 *
 * @param series    A series object created via `chart_t::add_series()`.
 * @param intercept Y-axis intercept value.
 *
 * The `%chart_series_set_trendline_intercept()` function sets the Y-axis
 * intercept for the trendline:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_equation(series);
 *  chart_series_set_trendline_intercept(series, 0.8);
 * @endcode
 *
 * @image html chart_trendline7.png
 *
 * As can be seen from the equation on the chart the intercept point
 * (when X=0) is the same as the value set in the equation.
 *
 * @note The intercept feature is only available in Excel for Exponential,
 * Linear and Polynomial trendline types.
 *
 * @see @ref chart_trendlines.
 */
void chart_series_set_trendline_intercept(chart_series_t& series, double intercept);

/**
 * @brief Set the trendline name for a chart data series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param name   The name of the trendline to display in the legend.
 *
 * The `%chart_series_set_trendline_name()` function sets the name of the
 * trendline that is displayed in the chart legend. In the examples above
 * the trendlines are displayed with default names like "Linear (Series 1)"
 * and "2 per Mov. Avg. (Series 1)". If these names are too verbose or not
 * descriptive enough you can set your own trendline name:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_name(series, "My trendline");
 * @endcode
 *
 * @image html chart_trendline8.png
 *
 * It is often preferable to turn off the trendline caption in the legend.
 * This is down in Excel by deleting the trendline name from the legend.
 * In `Xlsxwriter++` this is done using the `chart_legend_delete_series()`
 * function to delete the zero based series numbers:
 *
 * @code
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *
 *  // Delete the series name for the second series (=1 in zero base).
 *  std::vector<int16_t> names{1};
 *  chart_legend_delete_series(chart, names);
 * @endcode
 *
 * @image html chart_trendline9.png
 *
 * @see @ref chart_trendlines.
 */
void chart_series_set_trendline_name(chart_series_t& series, std::string_view name);

/**
 * @brief Set the trendline line properties for a chart data series.
 *
 * @param series A series object created via `chart_t::add_series()`.
 * @param line   A `chart_line_t` struct.
 *
 * The `%chart_series_set_trendline_line()` function is used to set the line
 * properties of a trendline:
 *
 * @code
 *  xwpp::chart_line_t line{
 *    .color_     = xwpp::color_t::RED,
 *    .dash_type_ = xwpp::chart_line_dash_type_t::LONG_DASH
 *  };
 *
 *  xwpp::series_set_trendline(series, xwpp::chart_trendline_type_t::LINEAR, 0);
 *  chart_series_set_trendline_line(series, &line);
 * @endcode
 *
 * @image html chart_trendline10.png
 *
 * @see @ref chart_trendlines and @ref chart_lines.
 */
void series_set_trendline_line(chart_series_t& series, const std::optional<chart_line_t>& line);

/**
 * @brief Set the direction (up, down or both) of the error bars for a chart
 * series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param direction  The bar direction: `chart_error_bar_direction_t`.
 *
 * The `%chart_series_set_error_bars_direction()` function sets the
 * direction of the error bars:
 *
 * @code
 *  chart.series_set_error_bars(series.y_error_bars_,
 *    xwpp::chart_error_bar_type_t::STD_ERROR, 0);
 *
 *  chart_series_set_error_bars_direction(series.y_error_bars_,
 *    xwpp::chart_error_bar_direction_t::PLUS);
 * @endcode
 *
 * @image html chart_error_bars2.png
 *
 * The valid directions are:
 *
 * - `%chart_error_bar_direction_t::BOTH`: Error bar extends in both directions. The default.
 * - `%chart_error_bar_direction_t::PLUS`: Error bar extends in positive direction.
 * - `%chart_error_bar_direction_t::MINUS`: Error bar extends in negative direction.
 *
 * @see @ref chart_error_bars.
 */
void chart_series_set_error_bars_direction(series_error_bars_t& error_bars, chart_error_bar_direction_t direction);

/**
 * @brief Set the end cap type for the error bars of a chart series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param endcap     The error bar end cap type: `chart_error_bar_cap_t`.
 *
 * The `%chart_series_set_error_bars_endcap()` function sets the end cap
 * type for the error bars:
 *
 * @code
 *  chart.series_set_error_bars(series.y_error_bars_,
 *    xwpp::chart_error_bar_type_t::STD_ERROR, 0);
 *
 *  chart_series_set_error_bars_endcap(series.y_error_bars_,
 *    xwpp::chart_error_bar_cap_tNO_CAP);
 * @endcode
 *
 * @image html chart_error_bars3.png
 *
 * The valid values are:
 *
 * - `%chart_error_bar_cap_t::END_CAP`: Flat end cap. The default.
 * - `%chart_error_bar_cap_t::NO_CAP`: No end cap.
 *
 * @see @ref chart_error_bars.
 */
void chart_series_set_error_bars_endcap(series_error_bars_t& error_bars, chart_error_bar_cap_t endcap);

/**
 * @brief Set the line properties for a chart series error bars.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param line       A `chart_line_t` struct.
 *
 * The `%chart_series_set_error_bars_line()` function sets the line
 * properties for the error bars:
 *
 * @code
 *  xwpp::chart_line_t line{
 *    .color_     = xwpp::color_t::RED,
 *    .dash_type_ = xwpp::chart_line_dash_type_t::DASH_ROUND_DOT,
 *  };
 *
 *  chart.series_set_error_bars(series.y_error_bars_,
 *    xwpp::chart_error_bar_type_t::STD_ERROR, 0);
 *
 *  chart_series_set_error_bars_line(series.y_error_bars_, line);
 * @endcode
 *
 * @image html chart_error_bars4.png
 *
 * @see @ref chart_lines and @ref chart_error_bars.
 */
void chart_series_set_error_bars_line(series_error_bars_t& error_bars, const std::optional<chart_line_t>& line);

/**
 * @brief Set the name caption of the an axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param name The name caption of the axis.
 *
 * The `%chart_axis_set_name()` function sets the name (also known as title or
 * caption) for an axis. It can be used for the X or Y axes. The name is
 * displayed below an X axis and to the side of a Y axis.
 *
 * @code
 *  chart_axis_set_name(chart->x_axis, "Earnings per Quarter");
 *  chart_axis_set_name(chart->y_axis, "US Dollars (Millions)");
 * @endcode
 *
 * @image html chart_axis_set_name.png
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *  chart_axis_set_name(chart->x_axis, "=Sheet1!$B$1");
 * @endcode
 *
 * @see The `chart_axis_set_name_range()` function to see how to set the
 * name formula programmatically.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name(chart_axis_t& axis, const std::string& name);

/**
 * @brief Set a chart axis name formula using row and column values.
 *
 * @param axis      A pointer to a chart `chart_axis_t` object.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row_num   The zero indexed row number of the range.
 * @param col_num   The zero indexed column number of the range.
 *
 * The `%chart_axis_set_name_range()` function can be used to set an axis name
 * range and is an alternative to using `chart_axis_set_name()` and a string
 * formula:
 *
 * @code
 *  chart_axis_set_name_range(chart->x_axis, "Sheet1", 1, 0);
 *  chart_axis_set_name_range(chart->y_axis, "Sheet1", 2, 0);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name_range(chart_axis_t& axis, const std::string& sheetname, row_num_t row_num, col_num_t col_num);

/**
 * @brief Set the manual position of the chart axis name.
 *
 * @param axis   A pointer to a chart `chart_axis_t` object.
 * @param layout A pointer to a chart `chart_layout_t` struct.
 *
 * This function is used to simulate setting the manual position of a chart
 * axis name.
 *
 * @see @ref chart_layout for more information.
 */
void chart_axis_set_name_layout(chart_axis_t& axis, const std::optional<chart_layout_t>& layout);

/**
 * @brief Set the font properties for a chart axis name.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param font A pointer to a chart `chart_font_t` font struct.
 *
 * The `%chart_axis_set_name_font()` function is used to set the font of an
 * axis name:
 *
 * @code
 *  xwpp::chart_font_t font{
 *    .bold_  = true,
 *    .color_ = xwpp::color_t::BLUE,
 *  };
 *
 *  chart_axis_set_name(chart->x_axis, "Yearly data");
 *  chart_axis_set_name_font(chart->x_axis, font);
 * @endcode
 *
 * @image html chart_axis_set_name_font.png
 *
 * @see @ref chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name_font(chart_axis_t& axis, const std::optional<chart_font_t>& font);

/**
 * @brief Set the font properties for the numbers of a chart axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param font A pointer to a chart `chart_font_t` font struct.
 *
 * The `%chart_axis_set_num_font()` function is used to set the font of the
 * numbers on an axis:
 *
 * @code
 *  xwpp::chart_font_t font{
 *    .bold_  = true,
 *    .color_ = xwpp::color_t::BLUE
 *  };
 *
 *  chart_axis_set_num_font(chart->x_axis, font1);
 * @endcode
 *
 * @image html chart_axis_set_num_font.png
 *
 * @see @ref chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_num_font(chart_axis_t& axis, const std::optional<chart_font_t>& font);

/**
 * @brief Set the number format for a chart axis.
 *
 * @param axis       A pointer to a chart `chart_axis_t` object.
 * @param num_format The number format string.
 *
 * The `%chart_axis_set_num_format()` function is used to set the format of
 * the numbers on an axis:
 *
 * @code
 *  chart_axis_set_num_format(chart->x_axis, "0.00%");
 *  chart_axis_set_num_format(chart->y_axis, "$#,##0.00");
 * @endcode
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `format_t::set_num_format()`.
 *
 * @image html chart_axis_num_format.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_num_format(chart_axis_t& axis, std::string_view num_format);

/**
 * @brief Set the line properties for a chart axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param line A `chart_line_t` struct.
 *
 * Set the line properties of a chart axis:
 *
 * @code
 *  // Hide the Y axis.
 *  xwpp::chart_line_t line{.none_ = true};
 *
 *  chart_axis_set_line(chart->y_axis, line);
 * @endcode
 *
 * @image html chart_axis_set_line.png
 *
 * @see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line);

/**
 * @brief Set the fill properties for a chart axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param fill A `chart_fill_t` struct.
 *
 * Set the fill properties of a chart axis:
 *
 * @code
 *  xwpp::chart_fill_t fill{.color_ = xwpp:color_t::YELLOW};
 *
 *  chart_axis_set_fill(chart->y_axis, fill);
 * @endcode
 *
 * @image html chart_axis_set_fill.png
 *
 * @see @ref chart_fills.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_fill(chart_axis_t& axis, const std::optional<chart_fill_t>& fill);

/**
 * @brief Set the pattern properties for a chart axis.
 *
 * @param axis    A pointer to a chart `chart_axis_t` object.
 * @param pattern A `chart_pattern_t` struct.
 *
 * Set the pattern properties of a chart axis:
 *
 * @code
 *  chart_axis_set_pattern(chart->y_axis, pattern);
 * @endcode
 *
 * @see `chart_pattern_type_t` and @ref chart_patterns.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_pattern(chart_axis_t& axis, const std::optional<chart_pattern_t>& pattern);

/**
 * @brief Reverse the order of the axis categories or values.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 *
 * Reverse the order of the axis categories or values:
 *
 * @code
 *  chart_axis_set_reverse(chart->x_axis);
 * @endcode
 *
 * @image html chart_reverse.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_reverse(chart_axis_t& axis);

/**
 * @brief Set the position that the axis will cross the opposite axis.
 *
 * @param axis  A pointer to a chart `chart_axis_t` object.
 * @param value The category or value that the axis crosses at.
 *
 * Set the position that the axis will cross the opposite axis:
 *
 * @code
 *  chart_axis_set_crossing(chart->x_axis, 3);
 *  chart_axis_set_crossing(chart->y_axis, 8);
 * @endcode
 *
 * @image html chart_crossing1.png
 *
 * If crossing is omitted (the default) the crossing will be set automatically
 * by Excel based on the chart data.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_crossing(chart_axis_t& axis, double value);

/**
 * @brief Set the opposite axis crossing position as the axis maximum.
 *
 * @param axis  A pointer to a chart `chart_axis_t` object.
 *
 * Set the position that the opposite axis will cross as the axis maximum.
 * The default axis crossing position is generally the axis minimum so this
 * function can be used to reverse the location of the axes without reversing
 * the number sequence:
 *
 * @code
 *  chart_axis_set_crossing_max(chart->x_axis);
 *  chart_axis_set_crossing_max(chart->y_axis);
 * @endcode
 *
 * @image html chart_crossing2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_crossing_max(chart_axis_t& axis);

/**
 * @brief Set the opposite axis crossing position as the axis minimum.
 *
 * @param axis  A pointer to a chart `chart_axis_t` object.
 *
 * Set the position that the opposite axis will cross as the axis minimum.
 * The default axis crossing position is generally the axis minimum so this
 * function can be used to reverse the location of the axes without reversing
 * the number sequence:
 *
 * @code
 *  chart_axis_set_crossing_min(chart->x_axis);
 *  chart_axis_set_crossing_min(chart->y_axis);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_crossing_min(chart_axis_t& axis);

/**
 * @brief Turn off/hide an axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 *
 * Turn off, hide, a chart axis:
 *
 * @code
 *  chart_axis_off(chart->x_axis);
 * @endcode
 *
 * @image html chart_axis_off.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_off(chart_axis_t& axis);

/**
 * @brief Position a category axis on or between the axis tick marks.
 *
 * @param axis     A pointer to a chart `chart_axis_t` object.
 * @param position A `chart_axis_tick_position_t` value.
 *
 * Position a category axis horizontally on, or between, the axis tick marks.
 *
 * There are two allowable values:
 *
 * - `%chart_axis_tick_position_t::ON_TICK`
 * - `%chart_axis_tick_position_t::BETWEEN`
 *
 * @code
 *  chart_axis_set_position(chart->x_axis, xwpp::chart_axis_tick_position_t::BETWEEN);
 * @endcode
 *
 * @image html chart_axis_set_position.png
 *
 * **Axis types**: This function is applicable to category axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_position(chart_axis_t& axis, chart_axis_tick_position_t position);

/**
 * @brief Position the axis labels.
 *
 * @param axis     A pointer to a chart `chart_axis_t` object.
 * @param position A `chart_axis_label_position_t` value.
 *
 * Position the axis labels for the chart. The labels are the numbers, or
 * strings or dates, on the axis that indicate the categories or values of
 * the axis.
 *
 * @code
 *  chart_axis_set_label_position(chart->x_axis, xwpp::chart_axis_label_position_t::HIGH);
 *  chart_axis_set_label_position(chart->y_axis, xwpp::chart_axis_label_position_t::HIGH);
 * @endcode
 *
 * @image html chart_label_position2.png
 *
 * The allowable values:
 *
 * - `%chart_axis_label_position_t::NEXT_TO` - The default.
 * - `%chart_axis_label_position_t::HIGH` - Also right for vertical axes.
 * - `%chart_axis_label_position_t::LOW` - Also left for vertical axes.
 * - `%chart_axis_label_position_t::NONE`
 *
 * @image html chart_label_position1.png
 *
 * The `chart_axis_label_position_t::NEXT_TO` turns off the axis labels. This
 * is slightly different from `chart_axis_off()` which also turns off the
 * labels but also turns off tick marks.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_label_position(chart_axis_t& axis, chart_axis_label_position_t position);

/**
 * @brief Set the alignment of the axis labels.
 *
 * @param axis  A pointer to a chart `chart_axis_t` object.
 * @param align A `chart_axis_label_alignment_t` value.
 *
 * Position the category axis labels for the chart. The labels are the
 * numbers, or strings or dates, on the axis that indicate the categories
 * of the axis.
 *
 * The allowable values:
 *
 * - `%chart_axis_label_alignment_t::CENTER` - Align label center (default).
 * - `%chart_axis_label_alignment_t::LEFT` - Align label left.
 * - `%chart_axis_label_alignment_t::RIGHT` - Align label right.
 *
 * @code
 *  chart_axis_set_label_align(chart->x_axis, xwpp::chart_axis_label_alignment_t::RIGHT);
 * @endcode
 *
 * **Axis types**: This function is applicable to category axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_label_align(chart_axis_t& axis, chart_axis_label_alignment_t align);

/**
 * @brief Set the minimum value for a chart axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param min  Minimum value for chart axis. Value axes only.
 *
 * Set the minimum value for the axis range.
 *
 * @code
 *  chart_axis_set_min(chart->y_axis, -4);
 *  chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * @image html chart_max_min.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_min(chart_axis_t& axis, double min);

/**
 * @brief Set the maximum value for a chart axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param max  Maximum value for chart axis. Value axes only.
 *
 * Set the maximum value for the axis range.
 *
 * @code
 *  chart_axis_set_min(chart->y_axis, -4);
 *  chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_max(chart_axis_t& axis, double max);

/**
 * @brief Set the log base of the axis range.
 *
 * @param axis     A pointer to a chart `chart_axis_t` object.
 * @param log_base The log base for value axis. Value axes only.
 *
 * Set the log base for the axis:
 *
 * @code
 *  chart_axis_set_log_base(chart->y_axis, 10);
 * @endcode
 *
 * @image html chart_log_base.png
 *
 * The allowable range of values for the log base in Excel is between 2 and
 * 1000.
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_log_base(chart_axis_t& axis, uint16_t log_base);

/**
 * @brief Set the major axis tick mark type.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param type The tick mark type, defined by `chart_tick_mark_t`.
 *
 * Set the type of the major axis tick mark:
 *
 * @code
 *  chart_axis_set_major_tick_mark(chart->x_axis, xwpp::chart_axis_tick_mark_t::CROSSING);
 *  chart_axis_set_minor_tick_mark(chart->x_axis, xwpp::chart_axis_tick_mark_t::INSIDE);
 *
 *  chart_axis_set_major_tick_mark(chart->x_axis, xwpp::chart_axis_tick_mark_t::OUTSIDE);
 *  chart_axis_set_minor_tick_mark(chart->y_axis, xwpp::chart_axis_tick_mark_t::INSIDE);
 *
 *  // Hide the default gridlines so the tick marks are visible.
 *  chart_axis_major_gridlines_set_visible(chart->y_axis, false);
 * @endcode
 *
 * @image html chart_tick_marks.png
 *
 * The tick mark types are:
 *
 * - `%chart_axis_tick_mark_t::NONE`
 * - `%chart_axis_tick_mark_t::INSIDE`
 * - `%chart_axis_tick_mark_t::OUTSIDE`
 * - `%chart_axis_tick_mark_t::CROSSING`
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_major_tick_mark(chart_axis_t& axis, chart_axis_tick_mark_t type);

/**
 * @brief Set the minor axis tick mark type.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param type The tick mark type, defined by `chart_tick_mark_t`.
 *
 * Set the type of the minor axis tick mark:
 *
 * @code
 *  chart_axis_set_minor_tick_mark(chart->x_axis, xwpp::chart_axis_tick_mark_t::INSIDE);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_minor_tick_mark(chart_axis_t& axis, chart_axis_tick_mark_t type);

/**
 * @brief Set the interval between category values.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param unit The interval between the categories.
 *
 * Set the interval between the category values. The default interval is 1
 * which gives the intervals shown in the charts above:
 *
 *     1, 2, 3, 4, 5, etc.
 *
 * Setting it to 2 gives:
 *
 *     1, 3, 5, 7, etc.
 *
 * @code
 *  chart_axis_set_interval_unit(chart->x_axis, 2);
 * @endcode
 *
 * @image html chart_set_interval1.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_interval_unit(chart_axis_t& axis, uint16_t unit);

/**
 * @brief Set the interval between category tick marks.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param unit The interval between the category ticks.
 *
 * Set the interval between the category tick marks. The default interval is 1
 * between each category but it can be set to other integer values:
 *
 * @code
 *  chart_axis_set_interval_tick(chart->x_axis, 2);
 * @endcode
 *
 * @image html chart_set_interval2.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_interval_tick(chart_axis_t& axis, uint16_t unit);

/**
 * @brief Set the increment of the major units in the axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param unit The increment of the major units.
 *
 * Set the increment of the major units in the axis range.
 *
 * @code
 *  // Turn on the minor gridline (it is off by default).
 *  chart_axis_minor_gridlines_set_visible(chart->y_axis, true);
 *
 *  chart_axis_set_major_unit(chart->y_axis, 4);
 *  chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * @image html chart_set_major_units.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_major_unit(chart_axis_t& axis, double unit);

/**
 * @brief Set the increment of the minor units in the axis.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param unit The increment of the minor units.
 *
 * Set the increment of the minor units in the axis range.
 *
 * @code
 *  chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * See the image above
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_minor_unit(chart_axis_t& axis, double unit);

/**
 * @brief Set the display units for a value axis.
 *
 * @param axis  A pointer to a chart `chart_axis_t` object.
 * @param units The display units: `chart_axis_display_unit_t`.
 *
 * Set the display units for the axis. This can be useful if the axis numbers
 * are very large but you don't want to represent them in scientific notation:
 *
 * @code
 *  chart_axis_set_display_units(chart->x_axis, xwpp::chart_axis_display_unit_t::THOUSANDS);
 *  chart_axis_set_display_units(chart->y_axis, xwpp::chart_axis_display_unit_t::MILLIONS);
 * @endcode
 *
 * @image html chart_display_units.png
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_display_units(chart_axis_t& axis, chart_axis_display_unit_t units);

/**
 * @brief Turn on/off the display units for a value axis.

 * @param axis    A pointer to a chart `chart_axis_t` object.
 * @param visible Turn off/on the display units. (0/1)
 *
 * Turn on or off the display units for the axis. This option is set on
 * automatically by `chart_axis_set_display_units()`.
 *
 * @code
 *  chart_axis_set_display_units_visible(chart->y_axis, true);
 * @endcode
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_display_units_visible(chart_axis_t& axis, bool visible);

/**
 * @brief Turn on/off the major gridlines for an axis.
 *
 * @param axis    A pointer to a chart `chart_axis_t` object.
 * @param visible Turn off/on the major gridline. (0/1)
 *
 * Turn on or off the major gridlines for an X or Y axis. In most Excel charts
 * the Y axis major gridlines are on by default and the X axis major
 * gridlines are off by default.
 *
 * @code
 *  // Reverse the normal visible/hidden gridlines for a column chart.
 *  chart_axis_major_gridlines_set_visible(chart->x_axis, true);
 *  chart_axis_major_gridlines_set_visible(chart->y_axis, false);
 * @endcode
 *
 * @image html chart_gridline1.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_major_gridlines_set_visible(chart_axis_t& axis, bool visible);

/**
 * @brief Turn on/off the minor gridlines for an axis.
 *
 * @param axis    A pointer to a chart `chart_axis_t` object.
 * @param visible Turn off/on the minor gridline. (0/1)
 *
 * Turn on or off the minor gridlines for an X or Y axis. In most Excel charts
 * the X and Y axis minor gridlines are off by default.
 *
 * Example, turn on all major and minor gridlines:
 *
 * @code
 *  chart_axis_major_gridlines_set_visible(chart->x_axis, true);
 *  chart_axis_minor_gridlines_set_visible(chart->x_axis, true);
 *  chart_axis_major_gridlines_set_visible(chart->y_axis, true);
 *  chart_axis_minor_gridlines_set_visible(chart->y_axis, true);
 * @endcode
 *
 * @image html chart_gridline2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_minor_gridlines_set_visible(chart_axis_t& axis, bool visible);

/**
 * @brief Set the line properties for the chart axis major gridlines.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param line A `chart_line_t` struct.
 *
 * Format the line properties of the major gridlines of a chart:
 *
 * @code
 *  xwpp::chart_line_t line1{
 *    .color_      = xwpp::color_t::RED,
 *    .width_      = 0.5,
 *    .dash_type_  = xwpp::xwpp::chart_line_dash_type_t::SQUARE_DOT,
 *  };
 *
 *  xwpp::chart_line_t line2{
 *    .color_      = xwpp::color_t::YELLOW,
 *  };
 *
 *  xwpp::chart_line_t line3{
 *    .width_      = 1.25,
 *    .dash_type   = xwpp::xwpp::chart_line_dash_type_t:DASH,
 *  };
 *
 *  xwpp::chart_line_t line4{
 *    .color_      = xwpp::color_t(0x00B050),
 *  };
 *
 *  chart_axis_major_gridlines_set_line(chart->x_axis, line1);
 *  chart_axis_minor_gridlines_set_line(chart->x_axis, line2);
 *  chart_axis_major_gridlines_set_line(chart->y_axis, line3);
 *  chart_axis_minor_gridlines_set_line(chart->y_axis, line4);
 * @endcode
 *
 * @image html chart_gridline3.png
 *
 * @see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_major_gridlines_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line);

/**
 * @brief Set the line properties for the chart axis minor gridlines.
 *
 * @param axis A pointer to a chart `chart_axis_t` object.
 * @param line A `chart_line_t` struct.
 *
 * Format the line properties of the minor gridlines of a chart, see the
 * example above.
 *
 * @see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_minor_gridlines_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line);

// TODO To rework (or remove). For test
void chart_add_data_cache(series_range_t& range, const uint8_t* data, uint16_t rows, uint8_t cols, uint8_t col);

/// @cond
}
/// @endcond

#endif
