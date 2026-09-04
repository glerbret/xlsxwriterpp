/*
 * chart - A library for creating Excel XLSX chart files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/chart.h"

#include "xwpp/exception.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <optional>

// TODO add gradient
namespace xwpp
{

namespace
{

std::optional<chart_font_t> convert_font_args(const std::optional<chart_font_t>& user_font, bool title_font = false)
{
  if(!user_font)
  {
    return std::nullopt;
  }

  chart_font_t font{
    .name_         = user_font->name_,
    .size_         = user_font->size_,
    .bold_         = user_font->bold_,
    .italic_       = user_font->italic_,
    .underline_    = user_font->underline_,
    .rotation_     = user_font->rotation_,
    .color_        = user_font->color_,
    .pitch_family_ = user_font->pitch_family_,
    .charset_      = user_font->charset_,
    .baseline_     = user_font->baseline_,
    .title_font_   = title_font,
  };

  // Convert font size units.
  if(font.size_ > 0.0)
  {
    font.size_ = font.size_ * 100.0;
  }

  // Convert rotation into 60,000ths of a degree.
  if(font.rotation_ != 0)
  {
    font.rotation_ = font.rotation_ * 60000;
  }

  return font;
}

std::optional<chart_line_t> convert_line_args(const std::optional<chart_line_t>& user_line)
{
  if(!user_line)
  {
    return std::nullopt;
  }

  chart_line_t line{
    .color_        = user_line->color_,
    .none_         = user_line->none_,
    .width_        = user_line->width_,
    .dash_type_    = user_line->dash_type_,
    .transparency_ = user_line->transparency_,
  };

  if(line.transparency_ > 100)
  {
    line.transparency_ = 0;
  }

  return line;
}

std::optional<chart_fill_t> convert_fill_args(const std::optional<chart_fill_t> user_fill)
{
  if(!user_fill)
  {
    return std::nullopt;
  }

  chart_fill_t fill{
    .color_        = user_fill->color_,
    .none_         = user_fill->none_,
    .transparency_ = user_fill->transparency_,
  };

  if(fill.transparency_ > 100)
  {
    fill.transparency_ = 0;
  }

  return fill;
}

std::optional<chart_pattern_t> convert_pattern_args(const std::optional<chart_pattern_t>& user_pattern)
{
  if(!user_pattern)
  {
    return std::nullopt;
  }

  if(user_pattern->type_ == chart_pattern_type_t::NONE)
  {
    return chart_pattern_t{};
  }

  if(user_pattern->fg_color_ == color_t::UNSET)
  {
    return chart_pattern_t{};
  }

  chart_pattern_t pattern{
    .fg_color_ = user_pattern->fg_color_,
    .bg_color_ = user_pattern->bg_color_,
    .type_     = user_pattern->type_,
  };

  if(pattern.bg_color_ == color_t::UNSET)
  {
    // Default background color in Excel is white, when unspecified.
    pattern.bg_color_ = color_t::WHITE;
  }

  return pattern;
}

std::optional<chart_layout_t> convert_layout_args(const std::optional<chart_layout_t>& user_layout,
                                                  chart_layout_type_t type)
{
  if(!user_layout)
  {
    return std::nullopt;
  }

  chart_layout_t layout;
  switch(type)
  {
    case chart_layout_type_t::LEGEND:
      layout.x_         = user_layout->x_;
      layout.y_         = user_layout->y_;
      layout.width_     = user_layout->width_;
      layout.height_    = user_layout->height_;
      layout.has_inner_ = false;
      break;

    case chart_layout_type_t::PLOTAREA:
      layout.x_         = user_layout->x_;
      layout.y_         = user_layout->y_;
      layout.width_     = user_layout->width_;
      layout.height_    = user_layout->height_;
      layout.has_inner_ = true;
      break;

    default:
      layout.x_         = user_layout->x_;
      layout.y_         = user_layout->y_;
      layout.width_     = 0.0;
      layout.height_    = 0.0;
      layout.has_inner_ = false;
      break;
  }

  return layout;
}

// Verify that a X/Y error bar property is supported for the chart type.
// All chart types, except Bar have Y error bars. Only Bar and Scatter
// support X error bars.
void check_error_bars(const series_error_bars_t& error_bars, std::string_view property)
{
  // Check that the error bar type has been set for all error bar
  // functions except the one that is used to set the type.
  if(!property.empty() && !error_bars.is_set_)
  {
    throw xwpp_exception_t(
      "check_error_bars(): error bar type must be set first using chart_t::series_set_error_bars().");
  }

  if(error_bars.is_x_)
  {
    if(error_bars.chart_group_ != chart_type_t::SCATTER && error_bars.chart_group_ != chart_type_t::BAR)
    {
      throw xwpp_exception_t(
        "check_error_bars(): 'X error bar' properties only available for Scatter and Bar charts in Excel.");
    }
  }
  else
  {
    if(error_bars.chart_group_ == chart_type_t::BAR)
    {
      throw xwpp_exception_t("check_error_bars(): 'Y error bar' properties not available for Bar charts in Excel.");
    }
  }
}

}

chart_t::chart_t(chart_type_t type)
  : type_{type}
{
  // Initialize the chart specific properties.
  initialize(type);
}

chart_series_t& chart_t::add_series(const std::string& categories, const std::string& values)
{
  // Scatter charts require categories and values.
  if(chart_group_ == chart_type_t::SCATTER && !values.empty() && categories.empty())
  {
    throw xwpp_exception_t("chart_t::add_series(): scatter charts must have 'categories' and 'values'.");
  }

  chart_series_t series;

  if(!categories.empty())
  {
    if(categories[0] == '=')
    {
      series.categories_.formula_ = categories.substr(1);
    }
    else
    {
      series.categories_.formula_ = categories;
    }
  }

  if(!values.empty())
  {
    if(values[0] == '=')
    {
      series.values_.formula_ = values.substr(1);
    }
    else
    {
      series.values_.formula_ = values;
    }
  }

  if(type_ == chart_type_t::SCATTER_SMOOTH)
  {
    series.smooth_ = true;
  }
  if(type_ == chart_type_t::SCATTER_SMOOTH_WITH_MARKERS)
  {
    series.smooth_ = true;
  }

  series.y_error_bars_.chart_group_ = chart_group_;
  series.x_error_bars_.chart_group_ = chart_group_;
  series.x_error_bars_.is_x_        = true;

  series.default_label_position_ = default_label_position_;

  series_list_.push_back(series);

  return series_list_.back();
}

void chart_t::set_style(uint8_t style_id)
{
  // The default style is 2. The range is 1 - 48
  if(style_id < 1 || style_id > 48)
  {
    style_id = 2;
  }

  style_id_ = style_id;
}

void chart_t::title_set_name(const std::string& name)
{
  if(!name.empty())
  {
    if(name[0] == '=')
    {
      title_.range_.formula_ = name.substr(1);
    }
    else
    {
      title_.name_ = name;
    }
  }
}

// cppcheck-suppress functionStatic
void chart_t::series_set_name(chart_series_t& series, const std::string& name) const
{
  if(!name.empty())
  {
    if(name[0] == '=')
    {
      series.title_.range_.formula_ = name.substr(1);
    }
    else
    {
      series.title_.name_ = name;
    }
  }
}

void chart_t::title_set_name_font(const chart_font_t& font)
{
  title_.font_ = convert_font_args(font, true);
}

void chart_t::legend_set_position(chart_legend_position_t position)
{
  legend_.position_ = position;
}

void chart_t::set_table()
{
  has_table_             = true;
  has_table_horizontal_  = true;
  has_table_vertical_    = true;
  has_table_outline_     = true;
  has_table_legend_keys_ = false;
}

void chart_t::set_table_grid(bool horizontal, bool vertical, bool outline, bool legend_keys)
{
  has_table_             = true;
  has_table_horizontal_  = horizontal;
  has_table_vertical_    = vertical;
  has_table_outline_     = outline;
  has_table_legend_keys_ = legend_keys;
}

void chart_t::set_drop_lines(const std::optional<chart_line_t>& line)
{
  has_drop_lines_  = true;
  drop_lines_line_ = convert_line_args(line);
}

void chart_t::set_high_low_lines(const std::optional<chart_line_t>& line)
{
  has_high_low_lines_  = true;
  high_low_lines_line_ = convert_line_args(line);
}

void chart_t::set_up_down_bars()
{
  has_up_down_bars_ = true;
}

void chart_t::set_up_down_bars_format(const std::optional<chart_line_t>& up_bar_line,
                                      const std::optional<chart_fill_t>& up_bar_fill,
                                      const std::optional<chart_line_t>& down_bar_line,
                                      const std::optional<chart_fill_t>& down_bar_fill)
{
  has_up_down_bars_ = true;
  up_bar_line_      = convert_line_args(up_bar_line);
  up_bar_fill_      = convert_fill_args(up_bar_fill);
  down_bar_line_    = convert_line_args(down_bar_line);
  down_bar_fill_    = convert_fill_args(down_bar_fill);
}

// cppcheck-suppress functionStatic
void chart_t::series_set_marker_type(chart_series_t& series, chart_marker_type_t type) const
{
  if(!series.marker_)
  {
    series.marker_ = chart_marker_t{};
  }

  series.marker_->type_ = type;
}

// cppcheck-suppress functionStatic
void chart_t::series_set_error_bars(series_error_bars_t& error_bars, chart_error_bar_type_t type, double value) const
{
  check_error_bars(error_bars, "");

  error_bars.type_      = type;
  error_bars.value_     = value;
  error_bars.has_value_ = true;
  error_bars.is_set_    = true;

  if(type == chart_error_bar_type_t::STD_ERROR)
  {
    error_bars.has_value_ = false;
  }
}

void chart_t::set_rotation(uint16_t rotation)
{
  if(rotation <= 360)
  {
    rotation_ = rotation;
  }
  else
  {
    throw xwpp_exception_t(
      std::format("chart_t::set_rotation(): chart rotation '{}' outside Excel range: 0 <= rotation <= 360.", rotation));
  }
}

void chart_t::set_hole_size(uint8_t size)
{
  if(size >= 10 && size <= 90)
  {
    hole_size_ = size;
  }
  else
  {
    throw xwpp_exception_t(
      std::format("chart_t::set_hole_size(): hole size '{}' outside Excel range: 10 <= size <= 90.", size));
  }
}

void chart_t::legend_set_font(const std::optional<chart_font_t>& font)
{
  legend_.font_ = convert_font_args(font);
}

void chart_t::set_series_gap(uint16_t gap)
{
  if(gap <= 500)
  {
    gap_y1_ = gap;
  }
  else
  {
    throw xwpp_exception_t(
      std::format("chart_t::set_series_gap(): chart series gap '{}' outside Excel range: 0 <= gap <= 500.", gap));
  }
}

void chart_t::chartarea_set_line(const std::optional<chart_line_t>& line)
{
  chartarea_line_ = convert_line_args(line);
}

void chart_t::chartarea_set_fill(const std::optional<chart_fill_t>& fill)
{
  chartarea_fill_ = convert_fill_args(fill);
}

void chart_t::plotarea_set_line(const std::optional<chart_line_t>& line)
{
  plotarea_line_ = convert_line_args(line);
}

void chart_t::plotarea_set_fill(const std::optional<chart_fill_t>& fill)
{
  plotarea_fill_ = convert_fill_args(fill);
}

void chart_t::show_blanks_as(chart_blank_t option)
{
  show_blanks_as_ = option;
}

void chart_t::show_hidden_data()
{
  show_hidden_data_ = true;
}

void chart_t::title_set_name_range(const std::string& sheetname, row_num_t row_num, col_num_t col_num)
{
  if(sheetname.empty())
  {
    throw xwpp_exception_t("chart_t::title_set_name_range(): sheetname must be specified.");
  }

  // Start and end row, col are the same for single cell range.
  set_range(title_.range_, sheetname, row_num, col_num, row_num, col_num);
}

void chart_t::legend_delete_series(const std::vector<int16_t>& delete_series)
{
  if(delete_series.empty())
  {
    throw xwpp_exception_t("chart_t::legend_delete_series(): 'delete_series' is empty.");
  }

  // The maximum number of series in a chart is 255.
  if(delete_series.size() > 255)
  {
    throw xwpp_exception_t("chart_t::legend_delete_series(): too manu elements in 'delete_series'.");
  }

  delete_series_ = delete_series;
}

void chart_t::set_series_overlap(int8_t overlap)
{
  if(overlap >= -100 && overlap <= 100)
  {
    overlap_y1_ = overlap;
  }
  else
  {
    throw xwpp_exception_t(std::format(
      "chart_t::set_series_overlap(): Chart series overlap '{}' outside Excel range: -100 <= overlap <= 100.",
      overlap));
  }
}

chart_axis_t& chart_t::axis_get(chart_axis_type_t axis_type)
{
  if(axis_type == chart_axis_type_t::TYPE_X)
  {
    return x_axis_;
  }
  else
  {
    return y_axis_;
  }
}

void chart_t::plotarea_set_layout(const std::optional<chart_layout_t>& layout)
{
  plotarea_layout_ = convert_layout_args(layout, chart_layout_type_t::PLOTAREA);
}

void chart_t::legend_set_layout(const std::optional<chart_layout_t>& layout)
{
  legend_.layout_ = convert_layout_args(layout, chart_layout_type_t::LEGEND);
}

void chart_t::title_set_layout(const std::optional<chart_layout_t>& layout)
{
  title_.layout_ = convert_layout_args(layout, chart_layout_type_t::TITLE);
}

void chart_t::title_set_overlay(bool overlay)
{
  title_.has_overlay_ = overlay;
}

void chart_t::plotarea_set_pattern(const std::optional<chart_pattern_t>& pattern)
{
  plotarea_pattern_ = convert_pattern_args(pattern);
}

void chart_t::set_table_font(const std::optional<chart_font_t>& font)
{
  has_table_  = true;
  table_font_ = convert_font_args(font);
}

void chart_t::title_off()
{
  title_.off_ = true;
}

// TODO Add test
void chart_t::chartarea_set_pattern(const std::optional<chart_pattern_t>& pattern)
{
  chartarea_pattern_ = convert_pattern_args(pattern);
}

std::string chart_t::assemble_xml_file()
{
  // Reverse the X and Y axes for Bar charts.
  if(type_ == chart_type_t::BAR || type_ == chart_type_t::BAR_STACKED || type_ == chart_type_t::BAR_STACKED_PERCENT)
  {
    std::swap(x_axis_, y_axis_);
  }

  std::string xml_data = xml_declaration();
  xml_data += write_chart_space();
  xml_data += write_lang();
  xml_data += write_style();
  if(is_protected_)
  {
    xml_data += write_protection();
  }
  xml_data += write_chart();
  xml_data += write_sp_pr(chartarea_line_, chartarea_fill_, chartarea_pattern_);
  if(!is_chartsheet_)
  {
    xml_data += write_print_settings();
  }
  xml_data += xml_end_tag("c:chartSpace");

  return xml_data;
}

void chart_t::set_axis_ids(uint32_t axis_id_1, uint32_t axis_id_2)
{
  axis_id_1_ = axis_id_1;
  axis_id_2_ = axis_id_2;
}

void chart_t::initialize_column_chart(chart_type_t type)
{
  chart_group_            = chart_type_t::COLUMN;
  has_horiz_val_axis_     = false;
  y_axis_.is_category_    = true;
  x_axis_.is_value_       = true;
  default_label_position_ = chart_label_position_t::OUTSIDE_END;

  if(type == chart_type_t::COLUMN_STACKED)
  {
    grouping_    = chart_grouping_t::STACKED;
    has_overlap_ = true;
    subtype_     = chart_subtype_t::STACKED;
    overlap_y1_  = 100;
  }

  if(type == chart_type_t::COLUMN_STACKED_PERCENT)
  {
    grouping_                   = chart_grouping_t::PERCENTSTACKED;
    y_axis_.default_num_format_ = "0%";
    has_overlap_                = true;
    subtype_                    = chart_subtype_t::STACKED;
    overlap_y1_                 = 100;
  }

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_column_chart;
  write_plot_area_  = write_plot_area;
}

void chart_t::initialize_bar_chart(chart_type_t type)
{
  // Note: Bar chart category/value axis are reversed in comparison to
  //       other charts. Some of the defaults reflect this.
  chart_group_                      = chart_type_t::BAR;
  x_axis_.major_gridlines_.visible_ = true;
  y_axis_.major_gridlines_.visible_ = false;
  y_axis_.is_category_              = true;
  x_axis_.is_value_                 = true;
  has_horiz_cat_axis_               = true;
  has_horiz_val_axis_               = false;
  default_label_position_           = chart_label_position_t::OUTSIDE_END;

  if(type == chart_type_t::BAR_STACKED)
  {
    grouping_    = chart_grouping_t::STACKED;
    has_overlap_ = true;
    subtype_     = chart_subtype_t::STACKED;
    overlap_y1_  = 100;
  }

  if(type == chart_type_t::BAR_STACKED_PERCENT)
  {
    grouping_                   = chart_grouping_t::PERCENTSTACKED;
    x_axis_.default_num_format_ = "0%";
    has_overlap_                = true;
    subtype_                    = chart_subtype_t::STACKED;
    overlap_y1_                 = 100;
  }

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_bar_chart;
  write_plot_area_  = write_plot_area;
}

void chart_t::initialize_area_chart(chart_type_t type)
{
  chart_group_            = chart_type_t::AREA;
  grouping_               = chart_grouping_t::STANDARD;
  default_cross_between_  = chart_axis_tick_position_t::ON_TICK;
  x_axis_.is_category_    = true;
  default_label_position_ = chart_label_position_t::CENTER;

  if(type == chart_type_t::AREA_STACKED)
  {
    grouping_ = chart_grouping_t::STACKED;
    subtype_  = chart_subtype_t::STACKED;
  }

  if(type == chart_type_t::AREA_STACKED_PERCENT)
  {
    grouping_                   = chart_grouping_t::PERCENTSTACKED;
    y_axis_.default_num_format_ = "0%";
    has_overlap_                = true;
    subtype_                    = chart_subtype_t::STACKED;
    overlap_y1_                 = 100;
  }

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_area_chart;
  write_plot_area_  = write_plot_area;
}

void chart_t::initialize_line_chart(chart_type_t type)
{
  chart_group_            = chart_type_t::LINE;
  default_marker_         = chart_marker_t{.type_ = chart_marker_type_t::NONE};
  grouping_               = chart_grouping_t::STANDARD;
  x_axis_.is_category_    = true;
  y_axis_.is_value_       = true;
  default_label_position_ = chart_label_position_t::RIGHT;

  if(type == chart_type_t::LINE_STACKED)
  {
    grouping_ = chart_grouping_t::STACKED;
    subtype_  = chart_subtype_t::STACKED;
  }

  if(type == chart_type_t::LINE_STACKED_PERCENT)
  {
    grouping_                   = chart_grouping_t::PERCENTSTACKED;
    y_axis_.default_num_format_ = "0%";
    subtype_                    = chart_subtype_t::STACKED;
  }

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_line_chart;
  write_plot_area_  = write_plot_area;
}

void chart_t::initialize_doughnut_chart()
{
  chart_group_            = chart_type_t::DOUGHNUT;
  write_chart_type_       = write_doughnut_chart;
  write_plot_area_        = write_pie_plot_area;
  default_label_position_ = chart_label_position_t::BEST_FIT;
}

void chart_t::initialize_pie_chart()
{
  chart_group_            = chart_type_t::PIE;
  write_chart_type_       = write_pie_chart;
  write_plot_area_        = write_pie_plot_area;
  default_label_position_ = chart_label_position_t::BEST_FIT;
}

void chart_t::initialize_radar_chart(chart_type_t type)
{
  if(type == chart_type_t::RADAR)
  {
    default_marker_ = chart_marker_t{.type_ = chart_marker_type_t::NONE};
  }

  chart_group_                      = chart_type_t::RADAR;
  x_axis_.major_gridlines_.visible_ = true;
  x_axis_.is_category_              = true;
  y_axis_.is_value_                 = true;
  y_axis_.major_tick_mark_          = chart_axis_tick_mark_t::CROSSING;
  default_label_position_           = chart_label_position_t::CENTER;

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_radar_chart;
  write_plot_area_  = write_plot_area;
}

void chart_t::initialize_scatter_chart(chart_type_t type)
{
  chart_group_            = chart_type_t::SCATTER;
  has_horiz_val_axis_     = false;
  default_cross_between_  = chart_axis_tick_position_t::ON_TICK;
  x_axis_.is_value_       = true;
  y_axis_.is_value_       = true;
  default_label_position_ = chart_label_position_t::RIGHT;

  if(type == chart_type_t::SCATTER_STRAIGHT || type == chart_type_t::SCATTER_SMOOTH)
  {
    default_marker_ = chart_marker_t{.type_ = chart_marker_type_t::NONE};
  }

  // Initialize the function pointers for this chart type.
  write_chart_type_ = write_scatter_chart;
  write_plot_area_  = write_scatter_plot_area;
}

// TODO Add stock chart
void chart_t::initialize(chart_type_t type)
{
  switch(type)
  {
    case chart_type_t::AREA:
    case chart_type_t::AREA_STACKED:
    case chart_type_t::AREA_STACKED_PERCENT:
      initialize_area_chart(type);
      break;

    case chart_type_t::BAR:
    case chart_type_t::BAR_STACKED:
    case chart_type_t::BAR_STACKED_PERCENT:
      initialize_bar_chart(type);
      break;

    case chart_type_t::COLUMN:
    case chart_type_t::COLUMN_STACKED:
    case chart_type_t::COLUMN_STACKED_PERCENT:
      initialize_column_chart(type);
      break;

    case chart_type_t::DOUGHNUT:
      initialize_doughnut_chart();
      break;

    case chart_type_t::LINE:
    case chart_type_t::LINE_STACKED:
    case chart_type_t::LINE_STACKED_PERCENT:
      initialize_line_chart(type);
      break;

    case chart_type_t::PIE:
      initialize_pie_chart();
      break;

    case chart_type_t::SCATTER:
    case chart_type_t::SCATTER_STRAIGHT:
    case chart_type_t::SCATTER_STRAIGHT_WITH_MARKERS:
    case chart_type_t::SCATTER_SMOOTH:
    case chart_type_t::SCATTER_SMOOTH_WITH_MARKERS:
      initialize_scatter_chart(type);
      break;

    case chart_type_t::RADAR:
    case chart_type_t::RADAR_WITH_MARKERS:
    case chart_type_t::RADAR_FILLED:
      initialize_radar_chart(type);
      break;

    default:
      throw xwpp_exception_t(
        std::format("chart_t::initialize(): unhandled chart type '{}'.", static_cast<uint16_t>(type)));
  }
}

std::string chart_t::write_bar_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:barChart");
  xml_data += write_bar_dir("bar");
  xml_data += write_grouping(chart.grouping_);
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_gap_width(chart.gap_y1_);
  xml_data += write_overlap(chart.overlap_y1_);
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:barChart");

  return xml_data;
}

std::string chart_t::write_column_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:barChart");
  xml_data += write_bar_dir("col");
  xml_data += write_grouping(chart.grouping_);
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_gap_width(chart.gap_y1_);
  xml_data += write_overlap(chart.overlap_y1_);
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:barChart");

  return xml_data;
}

std::string chart_t::write_plot_area(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:plotArea");
  xml_data += write_layout(chart.plotarea_layout_);
  xml_data += chart.write_chart_type_(chart);
  adjust_max_crossing(chart);
  xml_data += write_cat_axis(chart);
  xml_data += write_val_axis(chart);
  xml_data += write_d_table(chart);
  xml_data += write_sp_pr(chart.plotarea_line_, chart.plotarea_fill_, chart.plotarea_pattern_);
  xml_data += xml_end_tag("c:plotArea");

  return xml_data;
}

std::string chart_t::write_area_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:areaChart");
  xml_data += write_grouping(chart.grouping_);
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_drop_lines(chart);
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:areaChart");

  return xml_data;
}

std::string chart_t::write_line_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:lineChart");
  xml_data += write_grouping(chart.grouping_);
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_drop_lines(chart);
  xml_data += write_hi_low_lines(chart);
  xml_data += write_up_down_bars(chart);
  xml_data += write_marker_value();
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:lineChart");

  return xml_data;
}

std::string chart_t::write_doughnut_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:doughnutChart");
  xml_data += write_vary_colors();
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_first_slice_ang(chart);
  xml_data += write_hole_size(chart);
  xml_data += xml_end_tag("c:doughnutChart");

  return xml_data;
}

std::string chart_t::write_pie_plot_area(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:plotArea");
  xml_data += write_layout(chart.plotarea_layout_);
  xml_data += chart.write_chart_type_(chart);
  xml_data += write_sp_pr(chart.plotarea_line_, chart.plotarea_fill_, chart.plotarea_pattern_);
  xml_data += xml_end_tag("c:plotArea");

  return xml_data;
}

std::string chart_t::write_pie_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:pieChart");
  xml_data += write_vary_colors();
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_first_slice_ang(chart);
  xml_data += xml_end_tag("c:pieChart");

  return xml_data;
}

std::string chart_t::write_radar_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:radarChart");
  xml_data += write_radar_style(chart);
  for(auto& series: chart.series_list_)
  {
    xml_data += write_ser(chart, series);
  }
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:radarChart");

  return xml_data;
}

std::string chart_t::write_scatter_plot_area(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:plotArea");
  xml_data += write_layout(chart.plotarea_layout_);
  xml_data += chart.write_chart_type_(chart);
  adjust_max_crossing(chart);
  xml_data += write_cat_val_axis(chart);
  chart.has_horiz_val_axis_ = true;
  xml_data += write_val_axis(chart);
  xml_data += write_sp_pr(chart.plotarea_line_, chart.plotarea_fill_, chart.plotarea_pattern_);
  xml_data += xml_end_tag("c:plotArea");

  return xml_data;
}

std::string chart_t::write_scatter_chart(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:scatterChart");
  xml_data += write_scatter_style(chart);
  for(auto& series: chart.series_list_)
  {
    // Add default scatter chart formatting to the series data unless
    // it has already been specified by the user
    if(chart.type_ == chart_type_t::SCATTER && !series.line_)
    {
      chart_line_t line{
        .color_        = static_cast<color_t>(0x000000),
        .none_         = true,
        .width_        = 2.25,
        .dash_type_    = chart_line_dash_type_t::SOLID,
        .transparency_ = 0,
      };
      series.line_ = convert_line_args(line);
    }
    xml_data += write_xval_ser(chart, series);
  }
  xml_data += write_axis_ids(chart);
  xml_data += xml_end_tag("c:scatterChart");

  return xml_data;
}

std::string chart_t::write_chart_space()
{
  return xml_start_tag("c:chartSpace", {
                                         {"xmlns:c", SCHEMA_DRAWING + "/chart"          },
                                         {"xmlns:a", SCHEMA_DRAWING + "/main"           },
                                         {"xmlns:r", SCHEMA_OFFICEDOC + "/relationships"},
  });
}

std::string chart_t::write_lang()
{
  return xml_empty_tag("c:lang", {
                                   {"val", "en-US"}
  });
}

std::string chart_t::write_style() const
{
  // Don"t write an element for the default style, 2.
  if(style_id_ == 2)
  {
    return "";
  }

  return xml_empty_tag("c:style", {
                                    {"val", std::to_string(style_id_)}
  });
}

std::string chart_t::write_protection()
{
  return xml_empty_tag("c:protection");
}

std::string chart_t::write_print_settings()
{
  std::string xml_data = xml_start_tag("c:printSettings");
  xml_data += write_header_footer();
  xml_data += write_page_margins();
  xml_data += write_page_setup();
  xml_data += xml_end_tag("c:printSettings");

  return xml_data;
}

std::string chart_t::write_header_footer()
{
  return xml_empty_tag("c:headerFooter");
}

std::string chart_t::write_page_margins()
{
  return xml_empty_tag("c:pageMargins", {
                                          {"b",      "0.75"},
                                          {"l",      "0.7" },
                                          {"r",      "0.7" },
                                          {"t",      "0.75"},
                                          {"header", "0.3" },
                                          {"footer", "0.3" },
  });
}

std::string chart_t::write_page_setup()
{
  return xml_empty_tag("c:pageSetup");
}

std::string chart_t::write_chart()
{
  std::string xml_data = xml_start_tag("c:chart");
  xml_data += write_chart_title();
  xml_data += write_plot_area_(*this);
  xml_data += write_legend();
  xml_data += write_plot_vis_only();
  xml_data += write_disp_blanks_as();
  xml_data += xml_end_tag("c:chart");

  return xml_data;
}

std::string chart_t::write_chart_title() const
{
  if(title_.off_)
  {
    return write_auto_title_deleted();
  }
  else
  {
    return write_title(title_);
  }
}

std::string chart_t::write_auto_title_deleted()
{
  return xml_empty_tag("c:autoTitleDeleted", {
                                               {"val", "1"}
  });
}

std::string chart_t::write_tx_pr_pie(bool is_horizontal, const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("c:txPr");
  xml_data += write_a_body_pr(font.has_value() ? font->rotation_ : 0, is_horizontal);
  xml_data += write_a_lst_style();
  xml_data += write_a_p_pie(font);
  xml_data += xml_end_tag("c:txPr");

  return xml_data;
}

std::string chart_t::write_a_p_pie(const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:p");
  xml_data += write_a_p_pr_pie(font);
  xml_data += write_a_end_para_rpr();
  xml_data += xml_end_tag("a:p");

  return xml_data;
}

std::string chart_t::write_a_p_pr_pie(const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:pPr", {
                                                  {"rtl", "0"}
  });
  xml_data += write_a_def_rpr(font);
  xml_data += xml_end_tag("a:pPr");

  return xml_data;
}

std::string chart_t::write_title(const chart_title_t& title)
{
  if(!title.name_.empty())
  {
    return write_title_rich(title);
  }
  else if(!title.range_.formula_.empty())
  {
    return write_title_formula(title);
  }

  return "";
}

std::string chart_t::write_title_rich(const chart_title_t& title)
{
  std::string xml_data = xml_start_tag("c:title");
  xml_data += write_tx_rich(title.name_, title.is_horizontal_, title.font_);
  xml_data += write_layout(title.layout_);
  if(title.has_overlay_)
  {
    xml_data += write_overlay();
  }
  xml_data += xml_end_tag("c:title");

  return xml_data;
}

std::string chart_t::write_tx_rich(const std::string& name, bool is_horizontal, const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("c:tx");
  xml_data += write_rich(name, font, is_horizontal, false);
  xml_data += xml_end_tag("c:tx");

  return xml_data;
}

std::string chart_t::write_rich(const std::string& name, const std::optional<chart_font_t>& font, bool is_horizontal,
                                bool ignore_rich_pr)
{
  std::string xml_data = xml_start_tag("c:rich");
  xml_data += write_a_body_pr(font.has_value() ? font->rotation_ : 0, is_horizontal);
  xml_data += write_a_lst_style();
  xml_data += write_a_p_rich(name, font, ignore_rich_pr);
  xml_data += xml_end_tag("c:rich");

  return xml_data;
}

std::string chart_t::write_a_body_pr(int32_t rotation, bool is_horizontal)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(rotation == 0 && is_horizontal)
  {
    rotation = -5400000;
  }

  if(rotation != 0)
  {
    if(rotation == 16200000)
    {
      // 270 deg/stacked angle.
      attributes.emplace_back("rot", "0");
      attributes.emplace_back("vert", "wordArtVert");
    }
    else if(rotation == 16260000)
    {
      // 271 deg/East Asian vertical.
      attributes.emplace_back("rot", "0");
      attributes.emplace_back("vert", "eaVert");
    }
    else if(rotation == 21600000)
    {
      // 360 deg = 0 for y axis.
      attributes.emplace_back("rot", "0");
      attributes.emplace_back("vert", "horz");
    }
    else
    {
      attributes.emplace_back("rot", std::to_string(rotation));
      attributes.emplace_back("vert", "horz");
    }
  }

  return xml_empty_tag("a:bodyPr", attributes);
}

std::string chart_t::write_a_lst_style()
{
  return xml_empty_tag("a:lstStyle");
}

std::string chart_t::write_a_p_rich(const std::string& name, const std::optional<chart_font_t>& font,
                                    bool ignore_rich_pr)
{
  std::string xml_data = xml_start_tag("a:p");
  if(!ignore_rich_pr)
  {
    xml_data += write_a_p_pr_rich(font);
  }
  xml_data += write_a_r(name, font);
  xml_data += xml_end_tag("a:p");

  return xml_data;
}

std::string chart_t::write_a_p_pr_rich(const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:pPr");
  xml_data += write_a_def_rpr(font);
  xml_data += xml_end_tag("a:pPr");

  return xml_data;
}

std::string chart_t::write_a_def_rpr(const std::optional<chart_font_t>& font)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  bool has_color{false};
  bool has_latin{false};

  if(font)
  {
    has_color                   = font->color_ != color_t::UNSET;
    has_latin                   = !font->name_.empty() || font->pitch_family_ != 0 || font->charset_ != 0;
    const bool use_font_default = !has_color && !has_latin && font->baseline_ != -1;

    // Set the font attributes.
    if(font->size_ > 0.0)
    {
      attributes.emplace_back("sz", std::format("{}", font->size_));
    }

    // To manage defautl value for tile
    if(use_font_default)
    {
      attributes.emplace_back("b", std::format("{:d}", font->bold_.has_value() ? font->bold_.value() : false));
    }
    else if(font->bold_.has_value())
    {
      attributes.emplace_back("b", std::format("{:d}", font->bold_.value()));
    }

    if(use_font_default || font->italic_)
    {
      attributes.emplace_back("i", std::format("{:d}", font->italic_));
    }

    if(font->underline_)
    {
      attributes.emplace_back("u", "sng");
    }

    if(font->baseline_ != -1)
    {
      attributes.emplace_back("baseline", std::to_string(font->baseline_));
    }
  }

  // There are sub-elements if the font name or color have changed.
  // NOLINTBEGIN(bugprone-unchecked-optional-access)
  if(has_latin || has_color)
  {
    std::string xml_data = xml_start_tag("a:defRPr", attributes);

    if(has_color)
    {
      xml_data += write_a_solid_fill(font->color_, 0);
    }

    if(has_latin)
    {
      // Free and reuse the attribute list for the latin attributes.
      attributes.clear();

      if(!font->name_.empty())
      {
        attributes.emplace_back("typeface", font->name_);
      }

      if(font->pitch_family_ != 0)
      {
        attributes.emplace_back("pitchFamily", std::to_string(font->pitch_family_));
      }

      if(font->pitch_family_ != 0 || font->charset_ != 0)
      {
        attributes.emplace_back("charset", std::to_string(font->charset_));
      }

      xml_data += xml_empty_tag("a:latin", attributes);
    }

    xml_data += xml_end_tag("a:defRPr");

    return xml_data;
  }
  // NOLINTEND(bugprone-unchecked-optional-access)
  else
  {
    return xml_empty_tag("a:defRPr", attributes);
  }
}

std::string chart_t::write_a_r(const std::string& name, const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:r");
  xml_data += write_a_r_pr(font);
  xml_data += write_a_t(name);
  xml_data += xml_end_tag("a:r");

  return xml_data;
}

std::string chart_t::write_a_r_pr(const std::optional<chart_font_t>& font)
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  bool has_color{false};
  bool has_latin{false};

  attributes.emplace_back("lang", "en-US");

  if(font)
  {
    has_color                   = font->color_ != color_t::UNSET;
    has_latin                   = !font->name_.empty() || font->pitch_family_ != 0 || font->charset_ != 0;
    const bool use_font_default = !has_color && !has_latin && font->baseline_ != -1;

    // Set the font attributes.
    if(font->size_ > 0.0)
    {
      attributes.emplace_back("sz", std::format("{}", font->size_));
    }

    // To manage defautl value for tile
    if(use_font_default)
    {
      attributes.emplace_back("b", std::format("{:d}", font->bold_.has_value() ? font->bold_.value() : false));
    }
    else if(font->bold_.has_value())
    {
      attributes.emplace_back("b", std::format("{:d}", font->bold_.value()));
    }

    if(use_font_default || font->italic_)
    {
      attributes.emplace_back("i", std::format("{:d}", font->italic_));
    }

    if(font->underline_)
    {
      attributes.emplace_back("u", "sng");
    }

    if(font->baseline_ != -1)
    {
      attributes.emplace_back("baseline", std::to_string(font->baseline_));
    }
  }

  // There are sub-elements if the font name or color have changed.
  // NOLINTBEGIN(bugprone-unchecked-optional-access)
  if(has_latin || has_color)
  {
    std::string xml_data = xml_start_tag("a:rPr", attributes);

    if(has_color)
    {
      xml_data += write_a_solid_fill(font->color_, 0);
    }

    if(has_latin)
    {
      // Free and reuse the attribute list for the latin attributes.
      attributes.clear();

      if(!font->name_.empty())
      {
        attributes.emplace_back("typeface", font->name_);
      }

      if(font->pitch_family_ != 0)
      {
        attributes.emplace_back("pitchFamily", std::to_string(font->pitch_family_));
      }

      if(font->pitch_family_ != 0 || font->charset_ != 0)
      {
        attributes.emplace_back("charset", std::to_string(font->charset_));
      }
      xml_data += xml_empty_tag("a:latin", attributes);
    }
    xml_data += xml_end_tag("a:rPr");

    return xml_data;
  }
  // NOLINTEND(bugprone-unchecked-optional-access)
  else
  {
    return xml_empty_tag("a:rPr", attributes);
  }
}

std::string chart_t::write_a_t(const std::string& name)
{
  return xml_data_element("a:t", name);
}

std::string chart_t::write_layout(const std::optional<chart_layout_t>& layout)
{
  if(!layout)
  {
    return xml_empty_tag("c:layout");
  }
  else
  {
    std::string xml_data = xml_start_tag("c:layout");
    xml_data += write_manual_layout(*layout);
    xml_data += xml_end_tag("c:layout");

    return xml_data;
  }
}

std::string chart_t::write_manual_layout(const chart_layout_t& layout)
{
  std::string xml_data = xml_start_tag("c:manualLayout");

  if(layout.has_inner_)
  {
    xml_data += write_layout_target();
  }
  xml_data += write_layout_mode("c:xMode");
  xml_data += write_layout_mode("c:yMode");
  xml_data += write_layout_dimension("c:x", layout.x_);
  xml_data += write_layout_dimension("c:y", layout.y_);
  if(layout.width_ > 0.0)
  {
    xml_data += write_layout_dimension("c:w", layout.width_);
  }
  if(layout.height_ > 0.0)
  {
    xml_data += write_layout_dimension("c:h", layout.height_);
  }
  xml_data += xml_end_tag("c:manualLayout");

  return xml_data;
}

std::string chart_t::write_overlay()
{
  return xml_empty_tag("c:overlay", {
                                      {"val", "1"}
  });
}

std::string chart_t::write_bar_dir(const std::string& type)
{
  return xml_empty_tag("c:barDir", {
                                     {"val", type}
  });
}

std::string chart_t::write_grouping(chart_grouping_t grouping)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(grouping == chart_grouping_t::STANDARD)
  {
    attributes.emplace_back("val", "standard");
  }
  else if(grouping == chart_grouping_t::PERCENTSTACKED)
  {
    attributes.emplace_back("val", "percentStacked");
  }
  else if(grouping == chart_grouping_t::STACKED)
  {
    attributes.emplace_back("val", "stacked");
  }
  else
  {
    attributes.emplace_back("val", "clustered");
  }

  return xml_empty_tag("c:grouping", attributes);
}

std::string chart_t::write_gap_width(uint16_t gap)
{
  if(gap == DEFAULT_GAP)
  {
    return "";
  }

  return xml_empty_tag("c:gapWidth", {
                                       {"val", std::to_string(gap)}
  });
}

std::string chart_t::write_overlap(int8_t overlap)
{
  if(overlap == 0)
  {
    return "";
  }

  return xml_empty_tag("c:overlap", {
                                      {"val", std::to_string(overlap)}
  });
}

std::string chart_t::write_axis_id(uint32_t axis_id)
{
  return xml_empty_tag("c:axId", {
                                   {"val", std::to_string(axis_id)}
  });
}

std::string chart_t::write_axis_ids(chart_t& chart)
{
  if(chart.axis_id_1_ == 0)
  {
    add_axis_ids(chart);
  }

  std::string xml_data = write_axis_id(chart.axis_id_1_);
  xml_data += write_axis_id(chart.axis_id_2_);

  return xml_data;
}

std::string chart_t::write_ser(chart_t& chart, chart_series_t& series)
{
  const uint16_t index = chart.series_index_++;

  std::string xml_data = xml_start_tag("c:ser");
  xml_data += write_idx(index);
  xml_data += write_order(index);
  xml_data += write_series_name(series);
  xml_data += write_sp_pr(series.line_, series.fill_, series.pattern_);
  xml_data += write_marker(chart, series.marker_);
  xml_data += write_invert_if_negative(series);
  xml_data += write_points(chart, series);
  xml_data += write_d_lbls(series);
  xml_data += write_trendline(series);
  xml_data += write_error_bars(series);
  xml_data += write_cat(chart, series);
  xml_data += write_val(series);
  if(chart.chart_group_ == chart_type_t::SCATTER || chart.chart_group_ == chart_type_t::LINE)
  {
    xml_data += write_smooth(series.smooth_);
  }
  xml_data += xml_end_tag("c:ser");

  return xml_data;
}

std::string chart_t::write_idx(uint16_t index)
{
  return xml_empty_tag("c:idx", {
                                  {"val", std::to_string(index)}
  });
}

std::string chart_t::write_order(uint16_t index)
{
  return xml_empty_tag("c:order", {
                                    {"val", std::to_string(index)}
  });
}

std::string chart_t::write_series_name(const chart_series_t& series)
{
  if(!series.title_.name_.empty())
  {
    return write_tx_value(series.title_.name_);
  }
  else if(!series.title_.range_.formula_.empty())
  {
    return write_tx_formula(series.title_);
  }

  return "";
}

std::string chart_t::write_tx_value(const std::string& name)
{
  std::string xml_data = xml_start_tag("c:tx");
  xml_data += write_v_str(name);
  xml_data += xml_end_tag("c:tx");

  return xml_data;
}

std::string chart_t::write_v_str(const std::string& str)
{
  return xml_data_element("c:v", str);
}

std::string chart_t::write_tx_formula(const chart_title_t& title)
{
  std::string xml_data = xml_start_tag("c:tx");
  xml_data += write_str_ref(title.range_);
  xml_data += xml_end_tag("c:tx");

  return xml_data;
}

std::string chart_t::write_str_ref(const series_range_t& range)
{
  std::string xml_data = xml_start_tag("c:strRef");
  xml_data += write_f(range.formula_);
  if(!range.data_cache_.empty())
  {
    xml_data += write_str_cache(range);
  }
  xml_data += xml_end_tag("c:strRef");

  return xml_data;
}

std::string chart_t::write_f(const std::string& formula)
{
  return xml_data_element("c:f", formula);
}

std::string chart_t::write_sp_pr(const std::optional<chart_line_t>& line, const std::optional<chart_fill_t>& fill,
                                 const std::optional<chart_pattern_t>& pattern)
{
  if(!line && !fill && !pattern)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:spPr");
  if(fill && !pattern)
  {
    if(fill->none_)
    {
      xml_data += write_a_no_fill();
    }
    else
    {
      xml_data += write_a_solid_fill(fill->color_, fill->transparency_);
    }
  }

  if(pattern)
  {
    xml_data += write_a_patt_fill(*pattern);
  }

  if(line)
  {
    xml_data += write_a_ln(*line);
  }
  xml_data += xml_end_tag("c:spPr");

  return xml_data;
}

std::string chart_t::write_a_no_fill()
{
  return xml_empty_tag("a:noFill");
}

std::string chart_t::write_a_solid_fill(color_t color, uint8_t transparency)
{
  std::string xml_data = xml_start_tag("a:solidFill");
  xml_data += write_a_srgb_clr(color, transparency);
  xml_data += xml_end_tag("a:solidFill");

  return xml_data;
}

std::string chart_t::write_a_srgb_clr(color_t color, uint8_t transparency)
{
  const std::vector<std::tuple<std::string, std::string>> attributes{
    {"val", std::format("{:06X}", static_cast<uint32_t>(color) & COLOR_MASK)}
  };

  if(transparency != 0)
  {
    std::string xml_data = xml_start_tag("a:srgbClr", attributes);
    xml_data += write_a_alpha(transparency);
    xml_data += xml_end_tag("a:srgbClr");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("a:srgbClr", attributes);
  }
}

std::string chart_t::write_a_alpha(uint8_t transparency)
{
  return xml_empty_tag("a:alpha", {
                                    {"val", std::to_string((100 - transparency) * 1000)}
  });
}

std::string chart_t::write_a_patt_fill(const chart_pattern_t& pattern)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(pattern.type_ == chart_pattern_type_t::NONE)
  {
    attributes.emplace_back("prst", "none");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_5)
  {
    attributes.emplace_back("prst", "pct5");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_10)
  {
    attributes.emplace_back("prst", "pct10");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_20)
  {
    attributes.emplace_back("prst", "pct20");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_25)
  {
    attributes.emplace_back("prst", "pct25");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_30)
  {
    attributes.emplace_back("prst", "pct30");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_40)
  {
    attributes.emplace_back("prst", "pct40");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_50)
  {
    attributes.emplace_back("prst", "pct50");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_60)
  {
    attributes.emplace_back("prst", "pct60");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_70)
  {
    attributes.emplace_back("prst", "pct70");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_75)
  {
    attributes.emplace_back("prst", "pct75");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_80)
  {
    attributes.emplace_back("prst", "pct80");
  }
  else if(pattern.type_ == chart_pattern_type_t::PERCENT_90)
  {
    attributes.emplace_back("prst", "pct90");
  }
  else if(pattern.type_ == chart_pattern_type_t::LIGHT_DOWNWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "ltDnDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::LIGHT_UPWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "ltUpDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::DARK_DOWNWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "dkDnDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::DARK_UPWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "dkUpDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::WIDE_DOWNWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "wdDnDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::WIDE_UPWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "wdUpDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::LIGHT_VERTICAL)
  {
    attributes.emplace_back("prst", "ltVert");
  }
  else if(pattern.type_ == chart_pattern_type_t::LIGHT_HORIZONTAL)
  {
    attributes.emplace_back("prst", "ltHorz");
  }
  else if(pattern.type_ == chart_pattern_type_t::NARROW_VERTICAL)
  {
    attributes.emplace_back("prst", "narVert");
  }
  else if(pattern.type_ == chart_pattern_type_t::NARROW_HORIZONTAL)
  {
    attributes.emplace_back("prst", "narHorz");
  }
  else if(pattern.type_ == chart_pattern_type_t::DARK_VERTICAL)
  {
    attributes.emplace_back("prst", "dkVert");
  }
  else if(pattern.type_ == chart_pattern_type_t::DARK_HORIZONTAL)
  {
    attributes.emplace_back("prst", "dkHorz");
  }
  else if(pattern.type_ == chart_pattern_type_t::DASHED_DOWNWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "dashDnDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::DASHED_UPWARD_DIAGONAL)
  {
    attributes.emplace_back("prst", "dashUpDiag");
  }
  else if(pattern.type_ == chart_pattern_type_t::DASHED_HORIZONTAL)
  {
    attributes.emplace_back("prst", "dashHorz");
  }
  else if(pattern.type_ == chart_pattern_type_t::DASHED_VERTICAL)
  {
    attributes.emplace_back("prst", "dashVert");
  }
  else if(pattern.type_ == chart_pattern_type_t::SMALL_CONFETTI)
  {
    attributes.emplace_back("prst", "smConfetti");
  }
  else if(pattern.type_ == chart_pattern_type_t::LARGE_CONFETTI)
  {
    attributes.emplace_back("prst", "lgConfetti");
  }
  else if(pattern.type_ == chart_pattern_type_t::ZIGZAG)
  {
    attributes.emplace_back("prst", "zigZag");
  }
  else if(pattern.type_ == chart_pattern_type_t::WAVE)
  {
    attributes.emplace_back("prst", "wave");
  }
  else if(pattern.type_ == chart_pattern_type_t::DIAGONAL_BRICK)
  {
    attributes.emplace_back("prst", "diagBrick");
  }
  else if(pattern.type_ == chart_pattern_type_t::HORIZONTAL_BRICK)
  {
    attributes.emplace_back("prst", "horzBrick");
  }
  else if(pattern.type_ == chart_pattern_type_t::WEAVE)
  {
    attributes.emplace_back("prst", "weave");
  }
  else if(pattern.type_ == chart_pattern_type_t::PLAID)
  {
    attributes.emplace_back("prst", "plaid");
  }
  else if(pattern.type_ == chart_pattern_type_t::DIVOT)
  {
    attributes.emplace_back("prst", "divot");
  }
  else if(pattern.type_ == chart_pattern_type_t::DOTTED_GRID)
  {
    attributes.emplace_back("prst", "dotGrid");
  }
  else if(pattern.type_ == chart_pattern_type_t::DOTTED_DIAMOND)
  {
    attributes.emplace_back("prst", "dotDmnd");
  }
  else if(pattern.type_ == chart_pattern_type_t::SHINGLE)
  {
    attributes.emplace_back("prst", "shingle");
  }
  else if(pattern.type_ == chart_pattern_type_t::TRELLIS)
  {
    attributes.emplace_back("prst", "trellis");
  }
  else if(pattern.type_ == chart_pattern_type_t::SPHERE)
  {
    attributes.emplace_back("prst", "sphere");
  }
  else if(pattern.type_ == chart_pattern_type_t::SMALL_GRID)
  {
    attributes.emplace_back("prst", "smGrid");
  }
  else if(pattern.type_ == chart_pattern_type_t::LARGE_GRID)
  {
    attributes.emplace_back("prst", "lgGrid");
  }
  else if(pattern.type_ == chart_pattern_type_t::SMALL_CHECK)
  {
    attributes.emplace_back("prst", "smCheck");
  }
  else if(pattern.type_ == chart_pattern_type_t::LARGE_CHECK)
  {
    attributes.emplace_back("prst", "lgCheck");
  }
  else if(pattern.type_ == chart_pattern_type_t::OUTLINED_DIAMOND)
  {
    attributes.emplace_back("prst", "openDmnd");
  }
  else if(pattern.type_ == chart_pattern_type_t::SOLID_DIAMOND)
  {
    attributes.emplace_back("prst", "solidDmnd");
  }
  else
  {
    attributes.emplace_back("prst", "percent_50");
  }

  std::string xml_data = xml_start_tag("a:pattFill", attributes);

  if(pattern.fg_color_ != color_t::UNSET)
  {
    xml_data += write_a_fg_clr(pattern.fg_color_);
  }

  if(pattern.bg_color_ != color_t::UNSET)
  {
    xml_data += write_a_bg_clr(pattern.bg_color_);
  }

  xml_data += xml_end_tag("a:pattFill");

  return xml_data;
}

std::string chart_t::write_a_fg_clr(color_t color)
{
  std::string xml_data = xml_start_tag("a:fgClr");
  xml_data += write_a_srgb_clr(color, 0);
  xml_data += xml_end_tag("a:fgClr");

  return xml_data;
}

std::string chart_t::write_a_bg_clr(color_t color)
{
  std::string xml_data = xml_start_tag("a:bgClr");
  xml_data += write_a_srgb_clr(color, 0);
  xml_data += xml_end_tag("a:bgClr");

  return xml_data;
}

std::string chart_t::write_a_ln(const chart_line_t& line)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  // Round width to nearest 0.25, like Excel.
  const auto width_flt = static_cast<uint32_t>((line.width_ + 0.125) * 4.0F) / 4.0;

  // Convert to internal units.
  const auto width_int = static_cast<uint32_t>(0.5 + (12700.0 * width_flt));

  if(line.width_ > 0.0)
  {
    attributes.emplace_back("w", std::to_string(width_int));
  }

  if(line.none_ || line.color_ != color_t::UNSET || line.dash_type_ != chart_line_dash_type_t::SOLID)
  {
    std::string xml_data = xml_start_tag("a:ln", attributes);
    if(line.none_)
    {
      xml_data += write_a_no_fill();
    }
    else if(line.color_ != color_t::UNSET)
    {
      xml_data += write_a_solid_fill(line.color_, line.transparency_);
    }

    if(line.dash_type_ != chart_line_dash_type_t::SOLID)
    {
      xml_data += write_a_prst_dash(line.dash_type_);
    }

    xml_data += xml_end_tag("a:ln");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("a:ln", attributes);
  }
}

std::string chart_t::write_a_prst_dash(chart_line_dash_type_t dash_type)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(dash_type == chart_line_dash_type_t::ROUND_DOT)
  {
    attributes.emplace_back("val", "sysDot");
  }
  else if(dash_type == chart_line_dash_type_t::SQUARE_DOT)
  {
    attributes.emplace_back("val", "sysDash");
  }
  else if(dash_type == chart_line_dash_type_t::DASH_DOT)
  {
    attributes.emplace_back("val", "dashDot");
  }
  else if(dash_type == chart_line_dash_type_t::LONG_DASH)
  {
    attributes.emplace_back("val", "lgDash");
  }
  else if(dash_type == chart_line_dash_type_t::LONG_DASH_DOT)
  {
    attributes.emplace_back("val", "lgDashDot");
  }
  else if(dash_type == chart_line_dash_type_t::LONG_DASH_DOT_DOT)
  {
    attributes.emplace_back("val", "lgDashDotDot");
  }
  else if(dash_type == chart_line_dash_type_t::DOT)
  {
    attributes.emplace_back("val", "dot");
  }
  else if(dash_type == chart_line_dash_type_t::SYSTEM_DASH_DOT)
  {
    attributes.emplace_back("val", "sysDashDot");
  }
  else if(dash_type == chart_line_dash_type_t::SYSTEM_DASH_DOT_DOT)
  {
    attributes.emplace_back("val", "sysDashDotDot");
  }
  else
  {
    attributes.emplace_back("val", "dash");
  }

  return xml_empty_tag("a:prstDash", attributes);
}

std::string chart_t::write_str_cache(const series_range_t& range)
{
  std::string xml_data = xml_start_tag("c:strCache");
  xml_data += write_pt_count(range.num_data_points_);
  for(uint16_t index{0}; const auto& data_point: range.data_cache_)
  {
    xml_data += write_pt(index, data_point);
    index++;
  }
  xml_data += xml_end_tag("c:strCache");

  return xml_data;
}

std::string chart_t::write_pt_count(uint16_t num_data_points)
{
  return xml_empty_tag("c:ptCount", {
                                      {"val", std::to_string(num_data_points)}
  });
}

std::string chart_t::write_pt(uint16_t index, const series_data_point_t& data_point)
{
  // Ignore chart points that have no data.
  if(!data_point.no_data_)
  {
    std::string xml_data = xml_start_tag("c:pt", {
                                                   {"idx", std::to_string(index)}
    });

    if(data_point.is_string_ && !data_point.str_.empty())
    {
      xml_data += write_v_str(data_point.str_);
    }
    else
    {
      xml_data += write_v_num(data_point.number_);
    }
    xml_data += xml_end_tag("c:pt");

    return xml_data;
  }

  return "";
}

std::string chart_t::write_v_num(double number)
{
  // TODO Complex format to have 20000000 and not 2e07. To check to have better choice
  return xml_data_element("c:v", std::format("{:.10g}", number));
}

std::string chart_t::write_cat(chart_t& chart, const chart_series_t& series)
{
  const bool has_string_cache = series.categories_.has_string_cache_;

  // Ignore <c:cat> elements for charts without category values.
  if(series.categories_.formula_.empty())
  {
    return "";
  }

  chart.cat_has_num_fmt_ = !has_string_cache;

  std::string xml_data = xml_start_tag("c:cat");
  xml_data += write_data_cache(series.categories_, has_string_cache);
  xml_data += xml_end_tag("c:cat");

  return xml_data;
}

std::string chart_t::write_data_cache(const series_range_t& range, bool has_string_cache)
{
  if(has_string_cache)
  {
    return write_str_ref(range);
  }
  else
  {
    return write_num_ref(range);
  }
}

std::string chart_t::write_num_ref(const series_range_t& range)
{
  std::string xml_data = xml_start_tag("c:numRef");
  xml_data += write_f(range.formula_);
  if(!range.data_cache_.empty())
  {
    xml_data += write_num_cache(range);
  }
  xml_data += xml_end_tag("c:numRef");

  return xml_data;
}

std::string chart_t::write_num_cache(const series_range_t& range)
{
  std::string xml_data = xml_start_tag("c:numCache");
  xml_data += write_format_code();
  xml_data += write_pt_count(range.num_data_points_);
  for(uint16_t index{0}; const auto& data_point: range.data_cache_)
  {
    xml_data += write_num_pt(index, data_point);
    index++;
  }
  xml_data += xml_end_tag("c:numCache");

  return xml_data;
}

std::string chart_t::write_format_code()
{
  return xml_data_element("c:formatCode", "General");
}

std::string chart_t::write_num_pt(uint16_t index, const series_data_point_t& data_point)
{
  // Ignore chart points that have no data.
  if(data_point.no_data_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:pt", {
                                                 {"idx", std::to_string(index)}
  });
  xml_data += write_v_num(data_point.number_);
  xml_data += xml_end_tag("c:pt");

  return xml_data;
}

std::string chart_t::write_val_axis(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:valAx");
  xml_data += write_axis_id(chart.axis_id_2_);
  xml_data += write_scaling(chart.y_axis_.reverse_, chart.y_axis_.has_min_, chart.y_axis_.min_, chart.y_axis_.has_max_,
                            chart.y_axis_.max_, chart.y_axis_.log_base_);
  if(chart.y_axis_.hidden_)
  {
    xml_data += write_delete();
  }
  xml_data += write_axis_pos(chart.y_axis_.axis_position_, chart.x_axis_.reverse_);
  xml_data += write_major_gridlines(chart.y_axis_);
  xml_data += write_minor_gridlines(chart.y_axis_);
  chart.y_axis_.title_.is_horizontal_ = chart.has_horiz_val_axis_;
  xml_data += write_title(chart.y_axis_.title_);
  xml_data += write_number_format(chart.y_axis_);
  xml_data += write_major_tick_mark(chart.y_axis_);
  xml_data += write_minor_tick_mark(chart.y_axis_);
  xml_data += write_tick_label_pos(chart.y_axis_);
  xml_data += write_sp_pr(chart.y_axis_.line_, chart.y_axis_.fill_, chart.y_axis_.pattern_);
  xml_data += write_axis_font(chart.y_axis_.num_font_);
  xml_data += write_cross_axis(chart.axis_id_1_);
  if(!chart.x_axis_.has_crossing_ || chart.x_axis_.crossing_min_ || chart.x_axis_.crossing_max_)
  {
    xml_data += write_crosses(chart.x_axis_);
  }
  else
  {
    xml_data += write_crosses_at(chart.x_axis_);
  }
  xml_data += write_cross_between(chart, chart.x_axis_.position_axis_);
  xml_data += write_major_unit(chart.y_axis_);
  xml_data += write_minor_unit(chart.y_axis_);
  xml_data += write_disp_units(chart.y_axis_);
  xml_data += xml_end_tag("c:valAx");

  return xml_data;
}

std::string chart_t::write_scaling(bool reverse, bool has_min, double min, bool has_max, double max, uint16_t log_base)
{
  std::string xml_data = xml_start_tag("c:scaling");
  xml_data += write_log_base(log_base);
  xml_data += write_orientation(reverse);

  if(has_max)
  {
    xml_data += write_max(max);
  }

  if(has_min)
  {
    xml_data += write_min(min);
  }

  xml_data += xml_end_tag("c:scaling");

  return xml_data;
}

std::string chart_t::write_max(double max)
{
  return xml_empty_tag("c:max", {
                                  {"val", std::format("{}", max)}
  });
}

std::string chart_t::write_min(double min)
{
  return xml_empty_tag("c:min", {
                                  {"val", std::format("{}", min)}
  });
}

std::string chart_t::write_log_base(uint16_t log_base)
{
  if(log_base == 0)
  {
    return "";
  }

  return xml_empty_tag("c:logBase", {
                                      {"val", std::to_string(log_base)}
  });
}

std::string chart_t::write_orientation(bool reverse)
{
  if(reverse)
  {
    return xml_empty_tag("c:orientation", {
                                            {"val", "maxMin"}
    });
  }
  else
  {
    return xml_empty_tag("c:orientation", {
                                            {"val", "minMax"}
    });
  }
}

std::string chart_t::write_val(const chart_series_t& series)
{
  std::string xml_data = xml_start_tag("c:val");
  // Write the data cache elements. The string_cache is set to false since
  // this should always be a number series. */
  xml_data += write_data_cache(series.values_, false);
  xml_data += xml_end_tag("c:val");

  return xml_data;
}

std::string chart_t::write_marker(const chart_t& chart, std::optional<chart_marker_t>& marker)
{
  // If there isn't a user defined marker use the default, if this chart
  //  type one. The default usually turns the marker off. */
  if(!marker)
  {
    marker = chart.default_marker_;
  }

  if(!marker)
  {
    return "";
  }

  if(marker->type_ == chart_marker_type_t::AUTOMATIC)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:marker");
  xml_data += write_symbol(marker->type_);
  if(marker->size_ != 0)
  {
    xml_data += write_marker_size(marker->size_);
  }
  xml_data += write_sp_pr(marker->line_, marker->fill_, marker->pattern_);
  xml_data += xml_end_tag("c:marker");

  return xml_data;
}

std::string chart_t::write_cat_axis(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:catAx");
  xml_data += write_axis_id(chart.axis_id_1_);
  // Write the c:scaling element. Note we can't set max, min, or log base
  // for a Category axis in Excel.
  xml_data += write_scaling(chart.x_axis_.reverse_, false, 0.0, false, 0.0, 0);
  if(chart.x_axis_.hidden_)
  {
    xml_data += write_delete();
  }
  xml_data += write_axis_pos(chart.x_axis_.axis_position_, chart.y_axis_.reverse_);
  xml_data += write_major_gridlines(chart.x_axis_);
  xml_data += write_minor_gridlines(chart.x_axis_);
  chart.x_axis_.title_.is_horizontal_ = chart.has_horiz_cat_axis_;
  xml_data += write_title(chart.x_axis_.title_);
  xml_data += write_cat_number_format(chart, chart.x_axis_);
  xml_data += write_major_tick_mark(chart.x_axis_);
  xml_data += write_minor_tick_mark(chart.x_axis_);
  xml_data += write_tick_label_pos(chart.x_axis_);
  xml_data += write_sp_pr(chart.x_axis_.line_, chart.x_axis_.fill_, chart.x_axis_.pattern_);
  xml_data += write_axis_font(chart.x_axis_.num_font_);
  xml_data += write_cross_axis(chart.axis_id_2_);
  if(!chart.y_axis_.has_crossing_ || chart.y_axis_.crossing_min_ || chart.y_axis_.crossing_max_)
  {
    xml_data += write_crosses(chart.y_axis_);
  }
  else
  {
    xml_data += write_crosses_at(chart.y_axis_);
  }
  xml_data += write_auto();
  xml_data += write_label_align(chart.x_axis_);
  xml_data += write_label_offset();
  xml_data += write_tick_label_skip(chart.x_axis_);
  xml_data += write_tick_mark_skip(chart.x_axis_);
  xml_data += xml_end_tag("c:catAx");

  return xml_data;
}

std::string chart_t::write_delete()
{
  return xml_empty_tag("c:delete", {
                                     {"val", "1"}
  });
}

std::string chart_t::write_axis_pos(chart_position_t position, bool reverse)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  switch(position)
  {
    case chart_position_t::RIGHT:
      if(reverse)
      {
        attributes.emplace_back("val", "l");
      }
      else
      {
        attributes.emplace_back("val", "r");
      }
      break;
    case chart_position_t::LEFT:
      if(reverse)
      {
        attributes.emplace_back("val", "r");
      }
      else
      {
        attributes.emplace_back("val", "l");
      }
      break;

    case chart_position_t::TOP:
      if(reverse)
      {
        attributes.emplace_back("val", "b");
      }
      else
      {
        attributes.emplace_back("val", "t");
      }
      break;
    case chart_position_t::BOTTOM:
      if(reverse)
      {
        attributes.emplace_back("val", "t");
      }
      else
      {
        attributes.emplace_back("val", "b");
      }
      break;
  }

  return xml_empty_tag("c:axPos", attributes);
}

std::string chart_t::write_title_formula(const chart_title_t& title)
{
  std::string xml_data = xml_start_tag("c:title");
  xml_data += write_tx_formula(title);
  xml_data += write_layout(title.layout_);
  if(title.has_overlay_)
  {
    xml_data += write_overlay();
  }
  xml_data += write_tx_pr(title.is_horizontal_, title.font_);
  xml_data += xml_end_tag("c:title");

  return xml_data;
}

std::string chart_t::write_tx_pr(bool is_horizontal, const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("c:txPr");
  xml_data += write_a_body_pr(font.has_value() ? font->rotation_ : 0, is_horizontal);
  xml_data += write_a_lst_style();
  xml_data += write_a_p_formula(font);
  xml_data += xml_end_tag("c:txPr");

  return xml_data;
}

std::string chart_t::write_a_p_formula(const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:p");
  xml_data += write_a_p_pr_formula(font);
  xml_data += write_a_end_para_rpr();
  xml_data += xml_end_tag("a:p");

  return xml_data;
}

std::string chart_t::write_a_p_pr_formula(const std::optional<chart_font_t>& font)
{
  std::string xml_data = xml_start_tag("a:pPr");
  xml_data += write_a_def_rpr(font);
  xml_data += xml_end_tag("a:pPr");

  return xml_data;
}

std::string chart_t::write_axis_font(const std::optional<chart_font_t>& font)
{
  if(!font)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:txPr");
  xml_data += write_a_body_pr(font->rotation_, false);
  xml_data += write_a_lst_style();
  xml_data += xml_start_tag("a:p");
  xml_data += write_a_p_pr_rich(font);
  xml_data += write_a_end_para_rpr();
  xml_data += xml_end_tag("a:p");
  xml_data += xml_end_tag("c:txPr");

  return xml_data;
}

std::string chart_t::write_major_gridlines(const chart_axis_t& axis)
{
  if(!axis.major_gridlines_.visible_)
  {
    return "";
  }

  if(axis.major_gridlines_.line_)
  {
    std::string xml_data = xml_start_tag("c:majorGridlines");
    xml_data += write_sp_pr(axis.major_gridlines_.line_, std::nullopt, std::nullopt);
    xml_data += xml_end_tag("c:majorGridlines");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:majorGridlines");
  }
}

std::string chart_t::write_minor_gridlines(const chart_axis_t& axis)
{
  if(!axis.minor_gridlines_.visible_)
  {
    return "";
  }

  if(axis.minor_gridlines_.line_)
  {
    std::string xml_data = xml_start_tag("c:minorGridlines");
    xml_data += write_sp_pr(axis.minor_gridlines_.line_, std::nullopt, std::nullopt);
    xml_data += xml_end_tag("c:minorGridlines");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:minorGridlines");
  }
}

std::string chart_t::write_cat_number_format(const chart_t& chart, const chart_axis_t& axis)
{
  std::string num_format;
  uint8_t source_linked{1};
  bool default_format{true};

  // Set the number format to the axis default if not set.
  if(!axis.num_format_.empty())
  {
    num_format = axis.num_format_;
  }
  else
  {
    num_format = axis.default_num_format_;
  }

  // Check if a user defined number format has been set.
  if(num_format != axis.default_num_format_)
  {
    source_linked  = 0;
    default_format = false;
  }

  // Allow override of sourceLinked.
  if(axis.source_linked_ != 0)
  {
    source_linked = 1;
  }

  // Skip if cat doesn't have a num format (unless it is non-default).
  if(!chart.cat_has_num_fmt_ && default_format)
  {
    return "";
  }

  return xml_empty_tag("c:numFmt", {
                                     {"formatCode",   num_format                   },
                                     {"sourceLinked", std::to_string(source_linked)}
  });
}

std::string chart_t::write_tick_label_pos(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.label_position_ == chart_axis_label_position_t::HIGH)
  {
    attributes.emplace_back("val", "high");
  }
  else if(axis.label_position_ == chart_axis_label_position_t::LOW)
  {
    attributes.emplace_back("val", "low");
  }
  else if(axis.label_position_ == chart_axis_label_position_t::NONE)
  {
    attributes.emplace_back("val", "none");
  }
  else
  {
    attributes.emplace_back("val", "nextTo");
  }

  return xml_empty_tag("c:tickLblPos", attributes);
}

std::string chart_t::write_cross_axis(uint32_t axis_id)
{
  return xml_empty_tag("c:crossAx", {
                                      {"val", std::to_string(axis_id)}
  });
}

std::string chart_t::write_crosses(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.crossing_min_)
  {
    attributes.emplace_back("val", "min");
  }
  else if(axis.crossing_max_)
  {
    attributes.emplace_back("val", "max");
  }
  else
  {
    attributes.emplace_back("val", "autoZero");
  }

  return xml_empty_tag("c:crosses", attributes);
}

std::string chart_t::write_crosses_at(const chart_axis_t& axis)
{
  return xml_empty_tag("c:crossesAt", {
                                        {"val", std::format("{}", axis.crossing_)}
  });
}

std::string chart_t::write_auto()
{
  return xml_empty_tag("c:auto", {
                                   {"val", "1"}
  });
}

std::string chart_t::write_label_align(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.label_align_ == chart_axis_label_alignment_t::LEFT)
  {
    attributes.emplace_back("val", "l");
  }
  else if(axis.label_align_ == chart_axis_label_alignment_t::RIGHT)
  {
    attributes.emplace_back("val", "r");
  }
  else
  {
    attributes.emplace_back("val", "ctr");
  }

  return xml_empty_tag("c:lblAlgn", attributes);
}

std::string chart_t::write_label_offset()
{
  return xml_empty_tag("c:lblOffset", {
                                        {"val", "100"}
  });
}

/*
 * Write the <c:numberFormat> element. Note: It is assumed that if a user
 * defined number format is supplied (i.e., non-default) then the sourceLinked
 * attribute is 0. The user can override this if required.
 */
std::string chart_t::write_number_format(const chart_axis_t& axis)
{
  std::string num_format;
  uint8_t source_linked{1};

  // Set the number format to the axis default if not set.
  if(!axis.num_format_.empty())
  {
    num_format = axis.num_format_;
  }
  else
  {
    num_format = axis.default_num_format_;
  }

  // Check if a user defined number format has been set.
  if(num_format != axis.default_num_format_)
  {
    source_linked = 0;
  }

  // Allow override of sourceLinked.
  if(axis.source_linked_ != 0)
  {
    source_linked = 1;
  }

  return xml_empty_tag("c:numFmt", {
                                     {"formatCode",   num_format                   },
                                     {"sourceLinked", std::to_string(source_linked)},
  });
}

std::string chart_t::write_cross_between(const chart_t& chart, chart_axis_tick_position_t position)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(position == chart_axis_tick_position_t::DEFAULT)
  {
    position = chart.default_cross_between_;
  }

  if(position == chart_axis_tick_position_t::ON_TICK)
  {
    attributes.emplace_back("val", "midCat");
  }
  else
  {
    attributes.emplace_back("val", "between");
  }

  return xml_empty_tag("c:crossBetween", attributes);
}

std::string chart_t::write_legend_pos(const std::string& position)
{
  return xml_empty_tag("c:legendPos", {
                                        {"val", position}
  });
}

std::string chart_t::write_legend_entry(size_t index)
{
  std::string xml_data = xml_start_tag("c:legendEntry");
  xml_data += write_idx(delete_series_[index]);
  xml_data += write_delete();
  xml_data += xml_end_tag("c:legendEntry");

  return xml_data;
}

std::string chart_t::write_legend()
{
  bool has_overlay{false};

  if(legend_.position_ == chart_legend_position_t::NONE)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:legend");
  switch(legend_.position_)
  {
    case chart_legend_position_t::LEFT:
      xml_data += write_legend_pos("l");
      break;
    case chart_legend_position_t::TOP:
      xml_data += write_legend_pos("t");
      break;
    case chart_legend_position_t::BOTTOM:
      xml_data += write_legend_pos("b");
      break;
    case chart_legend_position_t::TOP_RIGHT:
      xml_data += write_legend_pos("tr");
      break;
    case chart_legend_position_t::OVERLAY_RIGHT:
      xml_data += write_legend_pos("r");
      has_overlay = true;
      break;
    case chart_legend_position_t::OVERLAY_LEFT:
      xml_data += write_legend_pos("l");
      has_overlay = true;
      break;
    case chart_legend_position_t::OVERLAY_TOP_RIGHT:
      xml_data += write_legend_pos("tr");
      has_overlay = true;
      break;
    default:
      xml_data += write_legend_pos("r");
  }

  for(size_t index{0}; index < delete_series_.size(); index++)
  {
    xml_data += write_legend_entry(index);
  }

  xml_data += write_layout(legend_.layout_);

  if(chart_group_ == chart_type_t::PIE || chart_group_ == chart_type_t::DOUGHNUT)
  {
    if(has_overlay)
    {
      xml_data += write_overlay();
    }
    xml_data += write_tx_pr_pie(false, legend_.font_);
  }
  else
  {
    if(legend_.font_)
    {
      xml_data += write_tx_pr(false, legend_.font_);
    }
    if(has_overlay)
    {
      xml_data += write_overlay();
    }
  }

  xml_data += xml_end_tag("c:legend");

  return xml_data;
}

std::string chart_t::write_plot_vis_only() const
{
  if(show_hidden_data_)
  {
    return "";
  }

  return xml_empty_tag("c:plotVisOnly", {
                                          {"val", "1"}
  });
}

std::string chart_t::write_drop_lines(const chart_t& chart)
{
  if(!chart.has_drop_lines_)
  {
    return "";
  }

  if(chart.drop_lines_line_)
  {
    std::string xml_data = xml_start_tag("c:dropLines");
    xml_data += write_sp_pr(chart.drop_lines_line_, std::nullopt, std::nullopt);
    xml_data += xml_end_tag("c:dropLines");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:dropLines");
  }
}

std::string chart_t::write_d_lbls(const chart_series_t& series)
{
  if(!series.has_labels_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:dLbls");

  if(!series.data_labels_.empty())
  {
    xml_data += write_custom_labels(series);
  }

  if(!series.label_num_format_.empty())
  {
    xml_data += write_label_num_fmt(series.label_num_format_);
  }

  xml_data += write_sp_pr(series.label_line_, series.label_fill_, series.label_pattern_);

  if(series.label_font_)
  {
    xml_data += write_tx_pr(false, series.label_font_);
  }

  if(series.label_position_ != chart_label_position_t::DEFAULT)
  {
    xml_data += write_d_lbl_pos(series.label_position_);
  }

  if(series.show_labels_legend_)
  {
    xml_data += write_show_legend_key();
  }

  if(series.show_labels_value_)
  {
    xml_data += write_show_val();
  }

  if(series.show_labels_category_)
  {
    xml_data += write_show_cat_name();
  }

  if(series.show_labels_name_)
  {
    xml_data += write_show_ser_name();
  }

  if(series.show_labels_percent_)
  {
    xml_data += write_show_percent();
  }

  if(series.label_separator_ != chart_label_separator_t::COMMA)
  {
    xml_data += write_separator(series.label_separator_);
  }

  if(series.show_labels_leader_)
  {
    xml_data += write_show_leader_lines();
  }

  xml_data += xml_end_tag("c:dLbls");

  return xml_data;
}

std::string chart_t::write_custom_label_str(const chart_series_t& series, const chart_custom_label_t& data_label)
{
  const bool ignore_rich_pr = !data_label.line_ && !data_label.fill_ && !data_label.pattern_;

  std::string xml_data = xml_empty_tag("c:layout");
  xml_data += xml_start_tag("c:tx");
  xml_data += write_rich(data_label.value_, data_label.font_, false, ignore_rich_pr);
  xml_data += xml_end_tag("c:tx");
  xml_data += write_sp_pr(data_label.line_, data_label.fill_, data_label.pattern_);

  if(series.label_position_ != chart_label_position_t::DEFAULT)
  {
    xml_data += write_d_lbl_pos(series.label_position_);
  }

  if(series.show_labels_value_)
  {
    xml_data += write_show_val();
  }

  if(series.show_labels_category_)
  {
    xml_data += write_show_cat_name();
  }

  if(series.show_labels_name_)
  {
    xml_data += write_show_ser_name();
  }

  return xml_data;
}

std::string chart_t::write_custom_labels(const chart_series_t& series)
{
  std::string xml_data;

  for(uint16_t index{0}; const auto& data_label: series.data_labels_)
  {
    if(!data_label.value_.empty() || data_label.range_ || data_label.hide_ || data_label.font_)
    {
      xml_data += xml_start_tag("c:dLbl");
      xml_data += write_idx(index);
      if(data_label.hide_)
      {
        xml_data += write_delete();
      }
      else if(!data_label.value_.empty())
      {
        xml_data += write_custom_label_str(series, data_label);
      }
      else if(data_label.range_.has_value())
      {
        xml_data += write_custom_label_formula(series, data_label);
      }
      else if(data_label.font_.has_value())
      {
        xml_data += write_custom_label_format_only(data_label);
      }
      xml_data += xml_end_tag("c:dLbl");
    }

    index++;
  }

  return xml_data;
}

std::string chart_t::write_d_lbl_pos(chart_label_position_t position)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(position == chart_label_position_t::RIGHT)
  {
    attributes.emplace_back("val", "r");
  }
  else if(position == chart_label_position_t::LEFT)
  {
    attributes.emplace_back("val", "l");
  }
  else if(position == chart_label_position_t::ABOVE)
  {
    attributes.emplace_back("val", "t");
  }
  else if(position == chart_label_position_t::BELOW)
  {
    attributes.emplace_back("val", "b");
  }
  else if(position == chart_label_position_t::INSIDE_BASE)
  {
    attributes.emplace_back("val", "inBase");
  }
  else if(position == chart_label_position_t::INSIDE_END)
  {
    attributes.emplace_back("val", "inEnd");
  }
  else if(position == chart_label_position_t::OUTSIDE_END)
  {
    attributes.emplace_back("val", "outEnd");
  }
  else if(position == chart_label_position_t::BEST_FIT)
  {
    attributes.emplace_back("val", "bestFit");
  }
  else
  {
    attributes.emplace_back("val", "ctr");
  }

  return xml_empty_tag("c:dLblPos", attributes);
}

std::string chart_t::write_show_val()
{
  return xml_empty_tag("c:showVal", {
                                      {"val", "1"}
  });
}

std::string chart_t::write_show_cat_name()
{
  return xml_empty_tag("c:showCatName", {
                                          {"val", "1"}
  });
}

std::string chart_t::write_show_ser_name()
{
  return xml_empty_tag("c:showSerName", {
                                          {"val", "1"}
  });
}

std::string chart_t::write_custom_label_formula(const chart_series_t& series, const chart_custom_label_t& data_label)
{
  std::string xml_data = xml_empty_tag("c:layout");
  xml_data += xml_start_tag("c:tx");
  // NOLINTNEXTLINE(bugprone-unchecked-optional-access)
  xml_data += write_str_ref(data_label.range_.value());
  xml_data += xml_end_tag("c:tx");
  xml_data += write_custom_label_format_only(data_label);

  if(series.label_position_ != chart_label_position_t::DEFAULT)
  {
    xml_data += write_d_lbl_pos(series.label_position_);
  }

  if(series.show_labels_value_)
  {
    xml_data += write_show_val();
  }

  if(series.show_labels_category_)
  {
    xml_data += write_show_cat_name();
  }

  if(series.show_labels_name_)
  {
    xml_data += write_show_ser_name();
  }

  return xml_data;
}

std::string chart_t::write_custom_label_format_only(const chart_custom_label_t& data_label)
{
  std::string xml_data;

  if(data_label.line_ || data_label.fill_ || data_label.pattern_)
  {
    xml_data += write_sp_pr(data_label.line_, data_label.fill_, data_label.pattern_);
    xml_data += write_tx_pr(false, data_label.font_);
  }
  else if(data_label.font_)
  {
    xml_data += xml_empty_tag("c:spPr");
    xml_data += write_tx_pr(false, data_label.font_);
  }

  return xml_data;
}

std::string chart_t::write_label_num_fmt(const std::string& format)
{
  return xml_empty_tag("c:numFmt", {
                                     {"formatCode",   format},
                                     {"sourceLinked", "0"   },
  });
}

std::string chart_t::write_show_legend_key()
{
  return xml_empty_tag("c:showLegendKey", {
                                            {"val", "1"}
  });
}

std::string chart_t::write_separator(chart_label_separator_t separator)
{
  if(separator == chart_label_separator_t::SEMICOLON)
  {
    return xml_data_element("c:separator", "; ");
  }
  else if(separator == chart_label_separator_t::PERIOD)
  {
    return xml_data_element("c:separator", ". ");
  }
  else if(separator == chart_label_separator_t::NEWLINE)
  {
    return xml_data_element("c:separator", "\n");
  }
  else if(separator == chart_label_separator_t::SPACE)
  {
    return xml_data_element("c:separator", " ");
  }
  else
  {
    return xml_data_element("c:separator", ", ");
  }
}

std::string chart_t::write_show_percent()
{
  return xml_empty_tag("c:showPercent", {
                                          {"val", "1"}
  });
}

std::string chart_t::write_show_leader_lines()
{
  return xml_empty_tag("c:showLeaderLines", {
                                              {"val", "1"}
  });
}

std::string chart_t::write_a_end_para_rpr()
{
  return xml_empty_tag("a:endParaRPr", {
                                         {"lang", "en-US"}
  });
}

std::string chart_t::write_d_table(const chart_t& chart)
{
  if(!chart.has_table_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:dTable");
  xml_data += write_show_horz_border(chart.has_table_horizontal_);
  xml_data += write_show_vert_border(chart.has_table_vertical_);
  xml_data += write_show_outline(chart.has_table_outline_);
  xml_data += write_show_keys(chart.has_table_legend_keys_);
  if(chart.table_font_)
  {
    xml_data += write_tx_pr(false, chart.table_font_);
  }
  xml_data += xml_end_tag("c:dTable");

  return xml_data;
}

std::string chart_t::write_show_horz_border(bool value)
{
  if(!value)
  {
    return "";
  }

  return xml_empty_tag("c:showHorzBorder", {
                                             {"val", "1"}
  });
}

std::string chart_t::write_show_vert_border(bool value)
{
  if(!value)
  {
    return "";
  }

  return xml_empty_tag("c:showVertBorder", {
                                             {"val", "1"}
  });
}

std::string chart_t::write_show_outline(bool value)
{
  if(!value)
  {
    return "";
  }

  return xml_empty_tag("c:showOutline", {
                                          {"val", "1"}
  });
}

std::string chart_t::write_show_keys(bool value)
{
  if(!value)
  {
    return "";
  }

  return xml_empty_tag("c:showKeys", {
                                       {"val", "1"}
  });
}

std::string chart_t::write_hi_low_lines(const chart_t& chart)
{
  if(!chart.has_high_low_lines_)
  {
    return "";
  }

  if(chart.high_low_lines_line_)
  {
    std::string xml_data = xml_start_tag("c:hiLowLines");
    xml_data += write_sp_pr(chart.high_low_lines_line_, std::nullopt, std::nullopt);
    xml_data += xml_end_tag("c:hiLowLines");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:hiLowLines");
  }
}

std::string chart_t::write_up_down_bars(const chart_t& chart)
{
  if(!chart.has_up_down_bars_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:upDownBars");
  xml_data += write_gap_width(150);
  xml_data += write_up_bars(chart.up_bar_line_, chart.up_bar_fill_);
  xml_data += write_down_bars(chart.down_bar_line_, chart.down_bar_fill_);
  xml_data += xml_end_tag("c:upDownBars");

  return xml_data;
}

std::string chart_t::write_marker_value()
{
  return xml_empty_tag("c:marker", {
                                     {"val", "1"}
  });
}

std::string chart_t::write_up_bars(const std::optional<chart_line_t>& line, const std::optional<chart_fill_t>& fill)
{
  if(line || fill)
  {
    std::string xml_data = xml_start_tag("c:upBars");
    xml_data += write_sp_pr(line, fill, std::nullopt);
    xml_data += xml_end_tag("c:upBars");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:upBars");
  }
}

std::string chart_t::write_down_bars(const std::optional<chart_line_t>& line, const std::optional<chart_fill_t>& fill)
{
  if(line || fill)
  {
    std::string xml_data = xml_start_tag("c:downBars");
    xml_data += write_sp_pr(line, fill, std::nullopt);
    xml_data += xml_end_tag("c:downBars");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("c:downBars");
  }
}

std::string chart_t::write_symbol(chart_marker_type_t type)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(type == chart_marker_type_t::SQUARE)
  {
    attributes.emplace_back("val", "square");
  }
  else if(type == chart_marker_type_t::DIAMOND)
  {
    attributes.emplace_back("val", "diamond");
  }
  else if(type == chart_marker_type_t::TRIANGLE)
  {
    attributes.emplace_back("val", "triangle");
  }
  else if(type == chart_marker_type_t::X)
  {
    attributes.emplace_back("val", "x");
  }
  else if(type == chart_marker_type_t::STAR)
  {
    attributes.emplace_back("val", "star");
  }
  else if(type == chart_marker_type_t::SHORT_DASH)
  {
    attributes.emplace_back("val", "short_dash");
  }
  else if(type == chart_marker_type_t::LONG_DASH)
  {
    attributes.emplace_back("val", "long_dash");
  }
  else if(type == chart_marker_type_t::CIRCLE)
  {
    attributes.emplace_back("val", "circle");
  }
  else if(type == chart_marker_type_t::PLUS)
  {
    attributes.emplace_back("val", "plus");
  }
  else
  {
    attributes.emplace_back("val", "none");
  }

  return xml_empty_tag("c:symbol", attributes);
}

std::string chart_t::write_marker_size(uint8_t size)
{
  return xml_empty_tag("c:size", {
                                   {"val", std::to_string(size)}
  });
}

std::string chart_t::write_error_bars(const chart_series_t& series)
{
  std::string xml_data = write_err_bars(series.x_error_bars_);
  xml_data += write_err_bars(series.y_error_bars_);

  return xml_data;
}

std::string chart_t::write_err_bars(const series_error_bars_t& error_bars)
{
  if(!error_bars.is_set_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:errBars");

  if(error_bars.chart_group_ != chart_type_t::BAR && error_bars.chart_group_ != chart_type_t::COLUMN)
  {
    xml_data += write_err_dir(error_bars.is_x_);
  }
  xml_data += write_err_bar_type(error_bars.direction_);
  xml_data += write_err_val_type(error_bars.type_);

  if(error_bars.endcap_ == chart_error_bar_cap_t::NO_CAP)
  {
    xml_data += write_no_end_cap();
  }

  if(error_bars.has_value_)
  {
    xml_data += write_error_val(error_bars.value_);
  }

  xml_data += write_sp_pr(error_bars.line_, std::nullopt, std::nullopt);
  xml_data += xml_end_tag("c:errBars");

  return xml_data;
}

std::string chart_t::write_err_dir(bool is_x)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(is_x)
  {
    attributes.emplace_back("val", "x");
  }
  else
  {
    attributes.emplace_back("val", "y");
  }

  return xml_empty_tag("c:errDir", attributes);
}

std::string chart_t::write_err_bar_type(chart_error_bar_direction_t direction)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(direction == chart_error_bar_direction_t::PLUS)
  {
    attributes.emplace_back("val", "plus");
  }
  else if(direction == chart_error_bar_direction_t::MINUS)
  {
    attributes.emplace_back("val", "minus");
  }
  else
  {
    attributes.emplace_back("val", "both");
  }

  return xml_empty_tag("c:errBarType", attributes);
}

std::string chart_t::write_err_val_type(chart_error_bar_type_t type)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(type == chart_error_bar_type_t::FIXED)
  {
    attributes.emplace_back("val", "fixedVal");
  }
  else if(type == chart_error_bar_type_t::PERCENTAGE)
  {
    attributes.emplace_back("val", "percentage");
  }
  else if(type == chart_error_bar_type_t::STD_DEV)
  {
    attributes.emplace_back("val", "stdDev");
  }
  else
  {
    attributes.emplace_back("val", "stdErr");
  }

  return xml_empty_tag("c:errValType", attributes);
}

std::string chart_t::write_no_end_cap()
{
  return xml_empty_tag("c:noEndCap", {
                                       {"val", "1"}
  });
}

std::string chart_t::write_error_val(double value)
{
  return xml_empty_tag("c:val", {
                                  {"val", std::format("{}", value)}
  });
}

std::string chart_t::write_trendline(const chart_series_t& series)
{
  if(!series.has_trendline_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:trendline");

  if(!series.trendline_name_.empty())
  {
    xml_data += write_name(series.trendline_name_);
  }

  xml_data += write_sp_pr(series.trendline_line_, std::nullopt, std::nullopt);
  xml_data += write_trendline_type(series.trendline_type_);

  if(series.trendline_type_ == chart_trendline_type_t::POLY && series.trendline_value_ >= 2)
  {
    xml_data += write_order(series.trendline_value_);
  }

  if(series.trendline_type_ == chart_trendline_type_t::AVERAGE && series.trendline_value_ >= 2)
  {
    xml_data += write_period(series.trendline_value_);
  }

  if(series.has_trendline_forecast_)
  {
    xml_data += write_forward(series.trendline_forward_);
    xml_data += write_backward(series.trendline_backward_);
  }

  if(series.has_trendline_intercept_)
  {
    xml_data += write_intercept(series.trendline_intercept_);
  }

  if(series.has_trendline_r_squared_)
  {
    xml_data += write_disp_rsqr();
  }

  if(series.has_trendline_equation_)
  {
    xml_data += write_disp_eq();
    xml_data += write_trendline_lbl();
  }

  xml_data += xml_end_tag("c:trendline");

  return xml_data;
}

std::string chart_t::write_name(const std::string& name)
{
  return xml_data_element("c:name", name);
}

std::string chart_t::write_trendline_type(chart_trendline_type_t type)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(type == chart_trendline_type_t::LOG)
  {
    attributes.emplace_back("val", "log");
  }
  else if(type == chart_trendline_type_t::POLY)
  {
    attributes.emplace_back("val", "poly");
  }
  else if(type == chart_trendline_type_t::POWER)
  {
    attributes.emplace_back("val", "power");
  }
  else if(type == chart_trendline_type_t::EXP)
  {
    attributes.emplace_back("val", "exp");
  }
  else if(type == chart_trendline_type_t::AVERAGE)
  {
    attributes.emplace_back("val", "movingAvg");
  }
  else
  {
    attributes.emplace_back("val", "linear");
  }

  return xml_empty_tag("c:trendlineType", attributes);
}

std::string chart_t::write_period(uint8_t value)
{
  return xml_empty_tag("c:period", {
                                     {"val", std::to_string(value)}
  });
}

std::string chart_t::write_forward(double value)
{
  return xml_empty_tag("c:forward", {
                                      {"val", std::format("{}", value)}
  });
}

std::string chart_t::write_backward(double value)
{
  return xml_empty_tag("c:backward", {
                                       {"val", std::format("{}", value)}
  });
}

std::string chart_t::write_intercept(double value)
{
  return xml_empty_tag("c:intercept", {
                                        {"val", std::format("{}", value)}
  });
}

std::string chart_t::write_disp_rsqr()
{
  return xml_empty_tag("c:dispRSqr", {
                                       {"val", "1"}
  });
}

std::string chart_t::write_disp_eq()
{
  return xml_empty_tag("c:dispEq", {
                                     {"val", "1"}
  });
}

std::string chart_t::write_trendline_lbl()
{
  std::string xml_data = xml_start_tag("c:trendlineLbl");
  xml_data += xml_empty_tag("c:layout");
  xml_data += xml_empty_tag("c:numFmt", {
                                          {"formatCode",   "General"},
                                          {"sourceLinked", "0"      },
  });
  xml_data += xml_end_tag("c:trendlineLbl");

  return xml_data;
}

std::string chart_t::write_invert_if_negative(const chart_series_t& series)
{
  if(!series.invert_if_negative_)
  {
    return "";
  }

  return xml_empty_tag("c:invertIfNegative", {
                                               {"val", "1"}
  });
}

std::string chart_t::write_points(const chart_t& chart, const chart_series_t& series)
{
  std::string xml_data;

  for(uint16_t index{0}; const auto& point: series.points_)
  {
    // Ignore empty points.
    if(point.line_ || point.fill_ || point.pattern_)
    {
      xml_data += write_d_pt(chart, point, index);
    }
    index++;
  }

  return xml_data;
}

std::string chart_t::write_d_pt(const chart_t& chart, const chart_point_t& point, uint16_t index)
{
  std::string xml_data = xml_start_tag("c:dPt");
  xml_data += write_idx(index);

  if(chart.chart_group_ == chart_type_t::SCATTER || chart.chart_group_ == chart_type_t::LINE)
  {
    xml_data += xml_start_tag("c:marker");
  }

  xml_data += write_sp_pr(point.line_, point.fill_, point.pattern_);

  if(chart.chart_group_ == chart_type_t::SCATTER || chart.chart_group_ == chart_type_t::LINE)
  {
    xml_data += xml_end_tag("c:marker");
  }

  xml_data += xml_end_tag("c:dPt");

  return xml_data;
}

std::string chart_t::write_vary_colors()
{
  return xml_empty_tag("c:varyColors", {
                                         {"val", "1"}
  });
}

std::string chart_t::write_first_slice_ang(const chart_t& chart)
{
  return xml_empty_tag("c:firstSliceAng", {
                                            {"val", std::to_string(chart.rotation_)}
  });
}

std::string chart_t::write_hole_size(const chart_t& chart)
{
  return xml_empty_tag("c:holeSize", {
                                       {"val", std::to_string(chart.hole_size_)}
  });
}

std::string chart_t::write_radar_style(const chart_t& chart)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(chart.type_ == chart_type_t::RADAR_FILLED)
  {
    attributes.emplace_back("val", "filled");
  }
  else
  {
    attributes.emplace_back("val", "marker");
  }

  return xml_empty_tag("c:radarStyle", attributes);
}

std::string chart_t::write_major_tick_mark(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.major_tick_mark_ == chart_axis_tick_mark_t::DEFAULT)
  {
    return "";
  }

  if(axis.major_tick_mark_ == chart_axis_tick_mark_t::NONE)
  {
    attributes.emplace_back("val", "none");
  }
  else if(axis.major_tick_mark_ == chart_axis_tick_mark_t::INSIDE)
  {
    attributes.emplace_back("val", "in");
  }
  else if(axis.major_tick_mark_ == chart_axis_tick_mark_t::CROSSING)
  {
    attributes.emplace_back("val", "cross");
  }
  else
  {
    attributes.emplace_back("val", "out");
  }

  return xml_empty_tag("c:majorTickMark", attributes);
}

std::string chart_t::write_minor_tick_mark(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.minor_tick_mark_ == chart_axis_tick_mark_t::DEFAULT)
  {
    return "";
  }

  if(axis.minor_tick_mark_ == chart_axis_tick_mark_t::NONE)
  {
    attributes.emplace_back("val", "none");
  }
  else if(axis.minor_tick_mark_ == chart_axis_tick_mark_t::INSIDE)
  {
    attributes.emplace_back("val", "in");
  }
  else if(axis.minor_tick_mark_ == chart_axis_tick_mark_t::CROSSING)
  {
    attributes.emplace_back("val", "cross");
  }
  else
  {
    attributes.emplace_back("val", "out");
  }

  return xml_empty_tag("c:minorTickMark", attributes);
}

std::string chart_t::write_cat_val_axis(chart_t& chart)
{
  std::string xml_data = xml_start_tag("c:valAx");
  xml_data += write_axis_id(chart.axis_id_1_);
  xml_data += write_scaling(chart.x_axis_.reverse_, chart.x_axis_.has_min_, chart.x_axis_.min_, chart.x_axis_.has_max_,
                            chart.x_axis_.max_, chart.x_axis_.log_base_);

  if(chart.x_axis_.hidden_)
  {
    xml_data += write_delete();
  }

  xml_data += write_axis_pos(chart.x_axis_.axis_position_, chart.y_axis_.reverse_);
  xml_data += write_major_gridlines(chart.x_axis_);
  xml_data += write_minor_gridlines(chart.x_axis_);
  chart.x_axis_.title_.is_horizontal_ = chart.has_horiz_val_axis_;
  xml_data += write_title(chart.x_axis_.title_);
  xml_data += write_number_format(chart.x_axis_);
  xml_data += write_major_tick_mark(chart.x_axis_);
  xml_data += write_minor_tick_mark(chart.x_axis_);
  xml_data += write_tick_label_pos(chart.x_axis_);
  xml_data += write_sp_pr(chart.x_axis_.line_, chart.x_axis_.fill_, chart.x_axis_.pattern_);
  xml_data += write_axis_font(chart.x_axis_.num_font_);
  xml_data += write_cross_axis(chart.axis_id_2_);

  if(!chart.y_axis_.has_crossing_ || chart.y_axis_.crossing_min_ || chart.y_axis_.crossing_max_)
  {
    xml_data += write_crosses(chart.y_axis_);
  }
  else
  {
    xml_data += write_crosses_at(chart.y_axis_);
  }

  xml_data += write_cross_between(chart, chart.y_axis_.position_axis_);
  xml_data += write_major_unit(chart.x_axis_);
  xml_data += write_minor_unit(chart.x_axis_);
  xml_data += write_disp_units(chart.x_axis_);
  xml_data += xml_end_tag("c:valAx");

  return xml_data;
}

std::string chart_t::write_xval_ser(chart_t& chart, chart_series_t& series)
{
  const uint16_t index = chart.series_index_++;

  std::string xml_data = xml_start_tag("c:ser");
  xml_data += write_idx(index);
  xml_data += write_order(index);
  xml_data += write_series_name(series);
  xml_data += write_sp_pr(series.line_, series.fill_, series.pattern_);
  xml_data += write_marker(chart, series.marker_);
  xml_data += write_points(chart, series);
  xml_data += write_d_lbls(series);
  xml_data += write_trendline(series);
  xml_data += write_error_bars(series);
  xml_data += write_x_val(series);
  xml_data += write_y_val(series);
  xml_data += write_smooth(series.smooth_);
  xml_data += xml_end_tag("c:ser");

  return xml_data;
}

std::string chart_t::write_x_val(const chart_series_t& series)
{
  const bool has_string_cache{series.categories_.has_string_cache_};

  std::string xml_data = xml_start_tag("c:xVal");
  xml_data += write_data_cache(series.categories_, has_string_cache);
  xml_data += xml_end_tag("c:xVal");

  return xml_data;
}

std::string chart_t::write_y_val(const chart_series_t& series)
{
  std::string xml_data = xml_start_tag("c:yVal");
  // Write the data cache elements. The string_cache is set to false since
  // this should always be a number series.
  xml_data += write_data_cache(series.values_, false);
  xml_data += xml_end_tag("c:yVal");

  return xml_data;
}

std::string chart_t::write_smooth(bool smooth)
{
  if(!smooth)
  {
    return "";
  }

  return xml_empty_tag("c:smooth", {
                                     {"val", "1"}
  });
}

std::string chart_t::write_major_unit(const chart_axis_t& axis)
{
  if(!axis.has_major_unit_)
  {
    return "";
  }

  return xml_empty_tag("c:majorUnit", {
                                        {"val", std::format("{}", axis.major_unit_)}
  });
}

std::string chart_t::write_minor_unit(const chart_axis_t& axis)
{
  if(!axis.has_minor_unit_)
  {
    return "";
  }

  return xml_empty_tag("c:minorUnit", {
                                        {"val", std::format("{}", axis.minor_unit_)}
  });
}

std::string chart_t::write_disp_units(const chart_axis_t& axis)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(axis.display_units_ == chart_axis_display_unit_t::NONE)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("c:dispUnits");
  switch(axis.display_units_)
  {
    case chart_axis_display_unit_t::THOUSANDS:
      attributes.emplace_back("val", "thousands");
      break;

    case chart_axis_display_unit_t::TEN_THOUSANDS:
      attributes.emplace_back("val", "tenThousands");
      break;

    case chart_axis_display_unit_t::HUNDRED_THOUSANDS:
      attributes.emplace_back("val", "hundredThousands");
      break;

    case chart_axis_display_unit_t::MILLIONS:
      attributes.emplace_back("val", "millions");
      break;

    case chart_axis_display_unit_t::TEN_MILLIONS:
      attributes.emplace_back("val", "tenMillions");
      break;

    case chart_axis_display_unit_t::HUNDRED_MILLIONS:
      attributes.emplace_back("val", "hundredMillions");
      break;

    case chart_axis_display_unit_t::BILLIONS:
      attributes.emplace_back("val", "billions");
      break;

    case chart_axis_display_unit_t::TRILLIONS:
      attributes.emplace_back("val", "trillions");
      break;

    case chart_axis_display_unit_t::HUNDREDS:
    default:
      attributes.emplace_back("val", "hundreds");
      break;
  }
  xml_data += xml_empty_tag("c:builtInUnit", attributes);

  if(axis.display_units_visible_)
  {
    xml_data += xml_start_tag("c:dispUnitsLbl");
    xml_data += xml_empty_tag("c:layout");
    xml_data += xml_end_tag("c:dispUnitsLbl");
  }

  xml_data += xml_end_tag("c:dispUnits");

  return xml_data;
}

std::string chart_t::write_scatter_style(const chart_t& chart)
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(chart.type_ == chart_type_t::SCATTER_SMOOTH || chart.type_ == chart_type_t::SCATTER_SMOOTH_WITH_MARKERS)
  {
    attributes.emplace_back("val", "smoothMarker");
  }
  else
  {
    attributes.emplace_back("val", "lineMarker");
  }

  return xml_empty_tag("c:scatterStyle", attributes);
}

std::string chart_t::write_tick_label_skip(const chart_axis_t& axis)
{
  if(axis.interval_unit_ == 0)
  {
    return "";
  }

  return xml_empty_tag("c:tickLblSkip", {
                                          {"val", std::to_string(axis.interval_unit_)}
  });
}

std::string chart_t::write_tick_mark_skip(const chart_axis_t& axis)
{
  if(axis.interval_tick_ == 0)
  {
    return "";
  }

  return xml_empty_tag("c:tickMarkSkip", {
                                           {"val", std::to_string(axis.interval_tick_)}
  });
}

std::string chart_t::write_disp_blanks_as() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(show_blanks_as_ != chart_blank_t::AS_ZERO && show_blanks_as_ != chart_blank_t::AS_CONNECTED)
  {
    return "";
  }

  if(show_blanks_as_ == chart_blank_t::AS_ZERO)
  {
    attributes.emplace_back("val", "zero");
  }
  else
  {
    attributes.emplace_back("val", "span");
  }

  return xml_empty_tag("c:dispBlanksAs", attributes);
}

std::string chart_t::write_layout_target()
{
  return xml_empty_tag("c:layoutTarget", {
                                           {"val", "inner"}
  });
}

std::string chart_t::write_layout_mode(const std::string& mode)
{
  return xml_empty_tag(mode, {
                               {"val", "edge"}
  });
}

std::string chart_t::write_layout_dimension(const std::string& dimension, double value)
{
  return xml_empty_tag(dimension, {
                                    {"val", std::format("{}", value)}
  });
}

void chart_t::add_axis_ids(chart_t& chart)
{
  const uint32_t chart_id{50010000 + chart.id_};
  const uint32_t axis_count{1};

  chart.axis_id_1_ = chart_id + axis_count;
  chart.axis_id_2_ = chart.axis_id_1_ + 1;
}

void chart_t::adjust_max_crossing(chart_t& chart)
{
  if(chart.x_axis_.crossing_max_)
  {
    if(chart.y_axis_.axis_position_ == chart_position_t::RIGHT)
    {
      chart.y_axis_.axis_position_ = chart_position_t::LEFT;
    }
    if(chart.y_axis_.axis_position_ == chart_position_t::LEFT)
    {
      chart.y_axis_.axis_position_ = chart_position_t::RIGHT;
    }
    if(chart.y_axis_.axis_position_ == chart_position_t::TOP)
    {
      chart.y_axis_.axis_position_ = chart_position_t::BOTTOM;
    }
    if(chart.y_axis_.axis_position_ == chart_position_t::BOTTOM)
    {
      chart.y_axis_.axis_position_ = chart_position_t::TOP;
    }
  }

  if(chart.y_axis_.crossing_max_)
  {
    if(chart.x_axis_.axis_position_ == chart_position_t::RIGHT)
    {
      chart.x_axis_.axis_position_ = chart_position_t::LEFT;
    }
    if(chart.x_axis_.axis_position_ == chart_position_t::LEFT)
    {
      chart.x_axis_.axis_position_ = chart_position_t::RIGHT;
    }
    if(chart.x_axis_.axis_position_ == chart_position_t::TOP)
    {
      chart.x_axis_.axis_position_ = chart_position_t::BOTTOM;
    }
    if(chart.x_axis_.axis_position_ == chart_position_t::BOTTOM)
    {
      chart.x_axis_.axis_position_ = chart_position_t::TOP;
    }
  }
}

void chart_series_set_categories(chart_series_t& series, const std::string& sheetname, row_num_t first_row,
                                 col_num_t first_col, row_num_t last_row, col_num_t last_col)
{
  if(sheetname.empty())
  {
    throw xwpp_exception_t("chart_series_set_categories(): sheetname must be specified.");
  }

  set_range(series.categories_, sheetname, first_row, first_col, last_row, last_col);
}

void chart_series_set_values(chart_series_t& series, const std::string& sheetname, row_num_t first_row,
                             col_num_t first_col, row_num_t last_row, col_num_t last_col)
{
  if(sheetname.empty())
  {
    throw xwpp_exception_t("chart_series_set_values(): sheetname must be specified.");
  }

  set_range(series.values_, sheetname, first_row, first_col, last_row, last_col);
}

// TODO series_range_t function
void set_range(series_range_t& range, const std::string& sheetname, row_num_t first_row, col_num_t first_col,
               row_num_t last_row, col_num_t last_col)
{
  range.sheetname_ = sheetname;
  range.first_row_ = first_row;
  range.first_col_ = first_col;
  range.last_row_  = last_row;
  range.last_col_  = last_col;
  range.formula_   = rowcol_to_formula_abs(sheetname, first_row, first_col, last_row, last_col);
}

void chart_series_set_name_range(chart_series_t& series, const std::string& sheetname, row_num_t row_num,
                                 col_num_t col_num)
{
  if(sheetname.empty())
  {
    throw xwpp_exception_t("chart_series_set_name_range(): sheetname must be specified.");
  }

  set_range(series.title_.range_, sheetname, row_num, col_num, row_num, col_num);
}

void chart_series_set_line(chart_series_t& series, const std::optional<chart_line_t>& line)
{
  series.line_ = convert_line_args(line);
}

void chart_series_set_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill)
{
  series.fill_ = convert_fill_args(fill);
}

void chart_series_set_invert_if_negative(chart_series_t& series)
{
  series.invert_if_negative_ = true;
}

void chart_series_set_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern)
{
  series.pattern_ = convert_pattern_args(pattern);
}

void chart_series_set_marker_size(chart_series_t& series, uint8_t size)
{
  if(size < 2 || size > 72)
  {
    throw xwpp_exception_t(
      std::format("chart_series_set_marker_size(): marker size '{}' outside Excel range: 2 <= size <= 72.", size));
  }

  if(!series.marker_)
  {
    series.marker_ = chart_marker_t{};
  }

  series.marker_->size_ = size;
}

void chart_series_set_marker_line(chart_series_t& series, const std::optional<chart_line_t>& line)
{
  if(!series.marker_)
  {
    series.marker_ = chart_marker_t{};
  }

  series.marker_->line_ = convert_line_args(line);
}

void chart_series_set_marker_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill)
{
  if(!series.marker_)
  {
    series.marker_ = chart_marker_t{};
  }

  series.marker_->fill_ = convert_fill_args(fill);
}

// TODO Add test
void chart_series_set_marker_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern)
{
  if(!series.marker_)
  {
    series.marker_ = chart_marker_t{};
  }

  series.marker_->pattern_ = convert_pattern_args(pattern);
}

void series_set_points(chart_series_t& series, const std::vector<chart_point_t>& points)
{
  if(points.empty())
  {
    throw xwpp_exception_t("series_set_points(): list of points shall not be empty.");
  }

  for(const auto& src_point: points)
  {
    const chart_point_t dst_point{
      .line_    = convert_line_args(src_point.line_),
      .fill_    = convert_fill_args(src_point.fill_),
      .pattern_ = convert_pattern_args(src_point.pattern_),
    };
    series.points_.push_back(dst_point);
  }
}

void chart_series_set_smooth(chart_series_t& series, bool smooth)
{
  series.smooth_ = smooth;
}

void chart_series_set_labels(chart_series_t& series)
{
  series.has_labels_        = true;
  series.show_labels_value_ = true;
}

void chart_series_set_labels_options(chart_series_t& series, bool show_name, bool show_category, bool show_value)
{
  series.has_labels_           = true;
  series.show_labels_name_     = show_name;
  series.show_labels_category_ = show_category;
  series.show_labels_value_    = show_value;
}

void chart_series_set_labels_custom(chart_series_t& series, const std::vector<chart_data_label_t>& data_labels)
{
  if(data_labels.empty())
  {
    throw xwpp_exception_t("chart_series_set_labels_custom(): list of labels shall not be empty.");
  }

  series.has_labels_ = true;

  // Set the Value label type if no other type is set.
  if(!series.show_labels_name_ && !series.show_labels_category_ && !series.show_labels_value_)
  {
    series.show_labels_value_ = true;
  }

  for(const auto& user_label: data_labels)
  {
    chart_custom_label_t data_label{
      .hide_    = user_label.hide_,
      .font_    = convert_font_args(user_label.font_),
      .line_    = convert_line_args(user_label.line_),
      .fill_    = convert_fill_args(user_label.fill_),
      .pattern_ = convert_pattern_args(user_label.pattern_),
    };

    if(!user_label.value_.empty())
    {
      if(user_label.value_[0] == '=')
      {
        // The value is a formula. Handle like other chart ranges.
        data_label.range_ = series_range_t{.formula_ = user_label.value_.substr(1)};
      }
      else
      {
        // The value is a simple string.
        data_label.value_ = user_label.value_;
      }
    }

    series.data_labels_.push_back(data_label);
  }
}

void chart_series_set_labels_separator(chart_series_t& series, chart_label_separator_t separator)
{
  series.has_labels_      = true;
  series.label_separator_ = separator;
}

void chart_series_set_labels_position(chart_series_t& series, chart_label_position_t position)
{
  series.has_labels_        = true;
  series.show_labels_value_ = true;

  if(position != series.default_label_position_)
  {
    series.label_position_ = position;
  }
}

void chart_series_set_labels_leader_line(chart_series_t& series)
{
  series.has_labels_         = true;
  series.show_labels_leader_ = true;
}

void chart_series_set_labels_legend(chart_series_t& series)
{
  series.has_labels_         = true;
  series.show_labels_legend_ = true;
}

void chart_series_set_labels_percentage(chart_series_t& series)
{
  series.has_labels_          = true;
  series.show_labels_percent_ = true;
}

void chart_series_set_labels_num_format(chart_series_t& series, std::string_view num_format)
{
  if(!num_format.empty())
  {
    series.label_num_format_ = num_format;
  }
}

void chart_series_set_labels_font(chart_series_t& series, const std::optional<chart_font_t>& font)
{
  series.label_font_ = convert_font_args(font);
}

void chart_series_set_labels_line(chart_series_t& series, const std::optional<chart_line_t>& line)
{
  series.label_line_ = convert_line_args(line);
}

void chart_series_set_labels_fill(chart_series_t& series, const std::optional<chart_fill_t>& fill)
{
  series.label_fill_ = convert_fill_args(fill);
}

void chart_series_set_labels_pattern(chart_series_t& series, const std::optional<chart_pattern_t>& pattern)
{
  series.label_pattern_ = convert_pattern_args(pattern);
}

void series_set_trendline(chart_series_t& series, chart_trendline_type_t type, uint8_t value)
{
  if(type == chart_trendline_type_t::POLY || type == chart_trendline_type_t::AVERAGE)
  {
    if(value < 2)
    {
      throw xwpp_exception_t(
        "series_set_trendline(): order/period value must be >= 2 for Polynomial and Moving Average types.");
    }
    series.trendline_value_type_ = type;
  }

  series.has_trendline_   = true;
  series.trendline_type_  = type;
  series.trendline_value_ = value;
}

void chart_series_set_trendline_forecast(chart_series_t& series, double forward, double backward)
{
  if(!series.has_trendline_)
  {
    throw xwpp_exception_t(
      "chart_series_set_trendline_forecast(): trendline type must be set first using series_set_trendline().");
  }

  if(series.trendline_type_ == chart_trendline_type_t::AVERAGE)
  {
    throw xwpp_exception_t(
      "chart_series_set_trendline_forecast(): forecast isn't available in Excel for a Moving Average trendline.");
  }

  series.has_trendline_forecast_ = true;
  series.trendline_forward_      = forward;
  series.trendline_backward_     = backward;
}

void chart_series_set_trendline_equation(chart_series_t& series)
{
  if(!series.has_trendline_)
  {
    throw xwpp_exception_t(
      "series_set_trendline_equation(): trendline type must be set first using series_set_trendline().");
  }

  if(series.trendline_type_ == chart_trendline_type_t::AVERAGE)
  {
    throw xwpp_exception_t(
      "series_set_trendline_equation(): equation isn't available in Excel for a Moving Average trendline.");
  }

  series.has_trendline_equation_ = true;
}

void chart_series_set_trendline_r_squared(chart_series_t& series)
{
  if(!series.has_trendline_)
  {
    throw xwpp_exception_t(
      "chart_series_set_trendline_r_squared(): trendline type must be set first using series_set_trendline().");
  }

  if(series.trendline_type_ == chart_trendline_type_t::AVERAGE)
  {
    throw xwpp_exception_t(
      "chart_series_set_trendline_r_squared(): R squared isn't available in Excel for a Moving Average trendline.");
  }

  series.has_trendline_r_squared_ = true;
}

void chart_series_set_trendline_intercept(chart_series_t& series, double intercept)
{
  if(!series.has_trendline_)
  {
    throw xwpp_exception_t(
      "chart_series_set_trendline_intercept(): trendline type must be set first using series_set_trendline().");
  }

  if(series.trendline_type_ != chart_trendline_type_t::EXP &&
     series.trendline_type_ != chart_trendline_type_t::LINEAR && series.trendline_type_ != chart_trendline_type_t::POLY)
  {
    throw xwpp_exception_t("chart_series_set_trendline_r_squared(): intercept is only available in Excel for "
                           "Exponential, Linear and Polynomial trendline types.");
  }

  series.has_trendline_intercept_ = true;
  series.trendline_intercept_     = intercept;
}

void chart_series_set_trendline_name(chart_series_t& series, std::string_view name)
{
  if(!name.empty())
  {
    series.trendline_name_ = name;
  }
}

void series_set_trendline_line(chart_series_t& series, const std::optional<chart_line_t>& line)
{
  series.trendline_line_ = convert_line_args(line);
}

void chart_series_set_error_bars_direction(series_error_bars_t& error_bars, chart_error_bar_direction_t direction)
{
  check_error_bars(error_bars, "_direction");

  error_bars.direction_ = direction;
}

void chart_series_set_error_bars_endcap(series_error_bars_t& error_bars, chart_error_bar_cap_t endcap)
{
  check_error_bars(error_bars, "_endcap");

  error_bars.endcap_ = endcap;
}

void chart_series_set_error_bars_line(series_error_bars_t& error_bars, const std::optional<chart_line_t>& line)
{
  check_error_bars(error_bars, "_line");

  error_bars.line_ = convert_line_args(line);
}

void chart_axis_set_name(chart_axis_t& axis, const std::string& name)
{
  if(name.empty())
  {
    return;
  }

  if(name[0] == '=')
  {
    axis.title_.range_.formula_ = name.substr(1);
  }
  else
  {
    axis.title_.name_ = name;
  }
}

void chart_axis_set_name_range(chart_axis_t& axis, const std::string& sheetname, row_num_t row_num, col_num_t col_num)
{
  if(sheetname.empty())
  {
    throw xwpp_exception_t("chart_axis_set_name_range(): sheetname must be specified.");
  }

  // Start and end row, col are the same for single cell range.
  set_range(axis.title_.range_, sheetname, row_num, col_num, row_num, col_num);
}

void chart_axis_set_name_layout(chart_axis_t& axis, const std::optional<chart_layout_t>& layout)
{
  axis.title_.layout_ = convert_layout_args(layout, chart_layout_type_t::AXIS_NAME);
}

void chart_axis_set_name_font(chart_axis_t& axis, const std::optional<chart_font_t>& font)
{
  if(!font)
  {
    return;
  }
  axis.title_.font_ = convert_font_args(font);
}

void chart_axis_set_num_font(chart_axis_t& axis, const std::optional<chart_font_t>& font)
{
  if(!font)
  {
    return;
  }
  axis.num_font_ = convert_font_args(font);
}

void chart_axis_set_num_format(chart_axis_t& axis, std::string_view num_format)
{
  axis.num_format_ = num_format;
}

void chart_axis_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line)
{
  if(!line)
  {
    return;
  }

  axis.line_ = convert_line_args(line);
}

void chart_axis_set_fill(chart_axis_t& axis, const std::optional<chart_fill_t>& fill)
{
  axis.fill_ = convert_fill_args(fill);
}

// TODO Add test
void chart_axis_set_pattern(chart_axis_t& axis, const std::optional<chart_pattern_t>& pattern)
{
  axis.pattern_ = convert_pattern_args(pattern);
}

void chart_axis_set_reverse(chart_axis_t& axis)
{
  axis.reverse_ = true;
}

void chart_axis_set_crossing(chart_axis_t& axis, double value)
{
  axis.has_crossing_ = true;
  axis.crossing_     = value;
}

void chart_axis_set_crossing_max(chart_axis_t& axis)
{
  axis.has_crossing_ = true;
  axis.crossing_max_ = true;
}

void chart_axis_set_crossing_min(chart_axis_t& axis)
{
  axis.has_crossing_ = true;
  axis.crossing_min_ = true;
}

// TODO Add test
void chart_axis_off(chart_axis_t& axis)
{
  axis.hidden_ = true;
}

void chart_axis_set_position(chart_axis_t& axis, chart_axis_tick_position_t position)
{
  axis.position_axis_ = position;
}

void chart_axis_set_label_position(chart_axis_t& axis, chart_axis_label_position_t position)
{
  axis.label_position_ = position;
}

void chart_axis_set_label_align(chart_axis_t& axis, chart_axis_label_alignment_t align)
{
  axis.label_align_ = align;
}

void chart_axis_set_min(chart_axis_t& axis, double min)
{
  axis.min_     = min;
  axis.has_min_ = true;
}

void chart_axis_set_max(chart_axis_t& axis, double max)
{
  axis.max_     = max;
  axis.has_max_ = true;
}

void chart_axis_set_log_base(chart_axis_t& axis, uint16_t log_base)
{
  // Excel log range is 2-1000.
  if(log_base >= 2 && log_base <= 1000)
  {
    axis.log_base_ = log_base;
  }
}

void chart_axis_set_major_tick_mark(chart_axis_t& axis, chart_axis_tick_mark_t type)
{
  axis.major_tick_mark_ = type;
}

void chart_axis_set_minor_tick_mark(chart_axis_t& axis, chart_axis_tick_mark_t type)
{
  axis.minor_tick_mark_ = type;
}

void chart_axis_set_interval_unit(chart_axis_t& axis, uint16_t unit)
{
  axis.interval_unit_ = unit;
}

void chart_axis_set_interval_tick(chart_axis_t& axis, uint16_t unit)
{
  axis.interval_tick_ = unit;
}

void chart_axis_set_major_unit(chart_axis_t& axis, double unit)
{
  axis.has_major_unit_ = true;
  axis.major_unit_     = unit;
}

void chart_axis_set_minor_unit(chart_axis_t& axis, double unit)
{
  axis.has_minor_unit_ = true;
  axis.minor_unit_     = unit;
}

void chart_axis_set_display_units(chart_axis_t& axis, chart_axis_display_unit_t units)
{
  axis.display_units_         = units;
  axis.display_units_visible_ = true;
}

void chart_axis_set_display_units_visible(chart_axis_t& axis, bool visible)
{
  axis.display_units_visible_ = visible;
}

void chart_axis_major_gridlines_set_visible(chart_axis_t& axis, bool visible)
{
  axis.major_gridlines_.visible_ = visible;
}

void chart_axis_minor_gridlines_set_visible(chart_axis_t& axis, bool visible)
{
  axis.minor_gridlines_.visible_ = visible;
}

void chart_axis_major_gridlines_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line)
{
  axis.major_gridlines_.line_ = convert_line_args(line);

  // If the gridline has a format it should also be visible.
  if(axis.major_gridlines_.line_)
  {
    axis.major_gridlines_.visible_ = true;
  }
}

void chart_axis_minor_gridlines_set_line(chart_axis_t& axis, const std::optional<chart_line_t>& line)
{
  axis.minor_gridlines_.line_ = convert_line_args(line);

  // If the gridline has a format it should also be visible.
  if(axis.minor_gridlines_.line_)
  {
    axis.minor_gridlines_.visible_ = true;
  }
}

void chart_add_data_cache(series_range_t& range, const uint8_t* data, uint16_t rows, uint8_t cols, uint8_t col)
{
  range.ignore_cache_    = true;
  range.num_data_points_ = rows;

  // Initialize the series range data cache.
  for(size_t i{0}; i < rows; i++)
  {
    // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-pointer-arithmetic)
    const series_data_point_t data_point{.number_ = static_cast<double>(data[(i * cols) + col])};
    range.data_cache_.push_back(data_point);
  }
}

}
