/*
 * chartsheet - A library for creating Excel XLSX chartsheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/chartsheet.h"

#include "xwpp/worksheet.h"
/// #include "xwpp/utility.h"
#include "xwpp/exception.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

chartsheet_t::chartsheet_t(const worksheet_init_data_t& init_data, std::function<int32_t(format_t*)> get_xf_index)
  : name_{init_data.name_}
  , quoted_name_{init_data.quoted_name_}
  , index_{init_data.index_}
  , active_sheet_{init_data.active_sheet_}
  , first_sheet_{init_data.first_sheet_}

{
  worksheet_.is_chartsheet_     = true;
  worksheet_.zoom_scale_normal_ = false;
  worksheet_.orientation_       = drawing_orientation_t::LANDSCAPE;
}

std::string chartsheet_t::get_sheet_name() const
{
  return name_;
}

uint16_t chartsheet_t::get_sheet_index() const
{
  return index_;
}

std::string chartsheet_t::write_chartsheet() const
{
  return xml_start_tag("chartsheet",
                       {
                           {"xmlns",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main"          },
                           {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
  });
}

std::string chartsheet_t::write_sheet_pr() const
{
  return worksheet_.write_sheet_pr();
}

std::string chartsheet_t::write_sheet_views()
{
  return worksheet_.write_sheet_views();
}

std::string chartsheet_t::write_page_margins() const
{
  return worksheet_.write_page_margins();
}

std::string chartsheet_t::write_drawings()
{
  return worksheet_.write_drawings();
}

std::string chartsheet_t::write_sheet_protection(const protection_obj_t& protection) const
{
  return worksheet_.write_sheet_protection(protection);
}

std::string chartsheet_t::write_page_setup() const
{
  return worksheet_.write_page_setup();
}

std::string chartsheet_t::write_header_footer() const
{
  return worksheet_.write_header_footer();
}

[[nodiscard]] std::string chartsheet_t::assemble_xml_file()
{
  std::string xml_data = xml_declaration();
  xml_data += write_chartsheet();
  xml_data += write_sheet_pr();
  xml_data += write_sheet_views();
  xml_data += write_sheet_protection(protection_);
  xml_data += write_page_margins();
  xml_data += write_page_setup();
  xml_data += write_header_footer();
  xml_data += write_drawings();
  xml_data += xml_end_tag("chartsheet");

  return xml_data;
}

void chartsheet_t::set_chart(chart_t* chart, const std::optional<chart_options_t>& user_options)
{
  if(!chart)
  {
    throw xwpp_exception_t("chartsheet_t::set_chart(): 'chart' must be present");
  }

  if(chart->in_use_)
  {
    throw xwpp_exception_t(
        "chartsheet_t::set_chart(): the same chart object cannot be inserted in a worksheet more than once");
  }

  if(chart->series_list_.empty())
  {
    throw xwpp_exception_t("chartsheet_t::set_chart(): chart must have a series");
  }

  object_properties_t object_props;
  if(user_options)
  {
    object_props.x_offset_ = user_options->x_offset_;
    object_props.y_offset_ = user_options->y_offset_;
    object_props.x_scale_  = user_options->x_scale_;
    object_props.y_scale_  = user_options->y_scale_;
  }

  object_props.width_  = 480;
  object_props.height_ = 288;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  // Store chart references so they can be ordered in the workbook.
  object_props.chart_ = chart;

  worksheet_.chart_data_.push_back(object_props);

  chart->in_use_        = true;
  chart->is_chartsheet_ = true;
  ///     chart->is_protected = self->is_protected;
  chart_                = *chart;
}

void chartsheet_t::set_chart(chart_t* chart)
{
  set_chart(chart, std::nullopt);
}

/// void
/// chartsheet_select(lxw_chartsheet *self)
/// {
///     self->selected = LXW_TRUE;
///
///     /* Selected worksheet can't be hidden. */
///     self->hidden = LXW_FALSE;
/// }

void chartsheet_t::activate()
{
  worksheet_.selected_ = true;
  worksheet_.active_   = true;

  // Active worksheet can't be hidden.
  worksheet_.hidden_ = false;

  *active_sheet_ = index_;
}

/// void
/// chartsheet_set_first_sheet(lxw_chartsheet *self)
/// {
///     /* Active worksheet can't be hidden. */
///     self->hidden = LXW_FALSE;
///
///     *self->first_sheet = self->index;
/// }

/// void
/// chartsheet_hide(lxw_chartsheet *self)
/// {
///     self->hidden = LXW_TRUE;
///
///     /* A hidden worksheet shouldn't be active or selected. */
///     self->selected = LXW_FALSE;
///
///     /* If this is active_sheet or first_sheet reset the workbook value. */
///     if (*self->first_sheet == self->index)
///         *self->first_sheet = 0;
///
///     if (*self->active_sheet == self->index)
///         *self->active_sheet = 0;
/// }

/// void
/// chartsheet_set_tab_color(lxw_chartsheet *self, lxw_color_t color)
/// {
///     self->worksheet->tab_color = color;
/// }

/// void
/// chartsheet_protect(lxw_chartsheet *self, const char *password,
///                    lxw_protection *options)
/// {
///     struct lxw_protection_obj *protect = &self->protection;
///
///     /* Copy any user parameters to the internal structure. */
///     if (options) {
///         protect->objects = options->no_objects;
///         protect->no_content = options->no_content;
///     }
///     else {
///         protect->objects = LXW_FALSE;
///         protect->no_content = LXW_FALSE;
///     }
///
///     if (password) {
///         uint16_t hash = lxw_hash_password(password);
///         lxw_snprintf(protect->hash, 5, "%X", hash);
///     }
///     else {
///         if (protect->objects && protect->no_content)
///             return;
///     }
///
///     protect->no_sheet = LXW_TRUE;
///     protect->scenarios = LXW_TRUE;
///     protect->is_configured = LXW_TRUE;
///
///     if (self->chart)
///         self->chart->is_protected = LXW_TRUE;
///     else
///         self->is_protected = LXW_TRUE;
/// }

/// void
/// chartsheet_set_zoom(lxw_chartsheet *self, uint16_t scale)
/// {
///     /* Confine the scale to Excel"s range */
///     if (scale < 10 || scale > 400) {
///         LXW_WARN("chartsheet_set_zoom(): "
///                  "Zoom factor scale outside range: 10 <= zoom <= 400.");
///         return;
///     }
///
///     self->worksheet->zoom = scale;
/// }

/// void
/// chartsheet_set_portrait(lxw_chartsheet *self)
/// {
///     worksheet_set_portrait(self->worksheet);
/// }

/// void
/// chartsheet_set_landscape(lxw_chartsheet *self)
/// {
///     worksheet_set_landscape(self->worksheet);
/// }

/// void
/// chartsheet_set_paper(lxw_chartsheet *self, uint8_t paper_size)
/// {
///     worksheet_set_paper(self->worksheet, paper_size);
/// }

/// void
/// chartsheet_set_margins(lxw_chartsheet *self, double left, double right,
///                        double top, double bottom)
/// {
///     worksheet_set_margins(self->worksheet, left, right, top, bottom);
/// }

/// lxw_error
/// chartsheet_set_header_opt(lxw_chartsheet *self, const char *string,
///                           lxw_header_footer_options *options)
/// {
///     return worksheet_set_header_opt(self->worksheet, string, options);
/// }

/// lxw_error
/// chartsheet_set_footer_opt(lxw_chartsheet *self, const char *string,
///                           lxw_header_footer_options *options)
/// {
///     return worksheet_set_footer_opt(self->worksheet, string, options);
/// }

/// lxw_error
/// chartsheet_set_header(lxw_chartsheet *self, const char *string)
/// {
///     return chartsheet_set_header_opt(self, string, NULL);
/// }

/// lxw_error
/// chartsheet_set_footer(lxw_chartsheet *self, const char *string)
/// {
///     return chartsheet_set_footer_opt(self, string, NULL);
/// }

}
