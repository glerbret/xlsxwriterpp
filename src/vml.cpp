/*****************************************************************************
 * vml - A library for creating Excel XLSX vml files.
 *
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/vml.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>
#include <tuple>
#include <vector>

namespace xwpp
{
vml_t::vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& comment_objs, uint32_t vml_shape_id,
             comment_display_t comment_display_default)
  : comment_objs_{comment_objs}
  , vml_data_id_str_{vml_data_id_str}
  , vml_shape_id_{vml_shape_id}
  , comment_display_default_{comment_display_default}
{
}

vml_t::vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& image_objs, uint32_t vml_shape_id)
  : image_objs_{image_objs}
  , vml_data_id_str_{vml_data_id_str}
  , vml_shape_id_{vml_shape_id}
{
}

std::string vml_t::write_visible() const
{
  return xml_empty_tag("x:Visible");
}

std::string vml_t::write_formula(const std::string& equation) const
{
  return xml_empty_tag("v:f", {
                                  {"eqn", equation}
  });
}

std::string vml_t::write_formulas() const
{
  std::string xml_data = xml_start_tag("v:formulas");
  xml_data += write_formula("if lineDrawn pixelLineWidth 0");
  xml_data += write_formula("sum @0 1 0");
  xml_data += write_formula("sum 0 0 @1");
  xml_data += write_formula("prod @2 1 2");
  xml_data += write_formula("prod @3 21600 pixelWidth");
  xml_data += write_formula("prod @3 21600 pixelHeight");
  xml_data += write_formula("sum @0 0 1");
  xml_data += write_formula("prod @6 1 2");
  xml_data += write_formula("prod @7 21600 pixelWidth");
  xml_data += write_formula("sum @8 21600 0");
  xml_data += write_formula("prod @7 21600 pixelHeight");
  xml_data += write_formula("sum @10 21600 0");
  xml_data += xml_end_tag("v:formulas");

  return xml_data;
}

/*
 * Write the <x:TextHAlign> element.
 */
/// STATIC void
/// _vml_write_text_halign(lxw_vml *self)
/// {
///
///     lxw_xml_data_element(self->file, "x:TextHAlign", "Center", NULL);
/// }

/*
 * Write the <x:TextVAlign> element.
 */
/// STATIC void
/// _vml_write_text_valign(lxw_vml *self)
/// {
///     lxw_xml_data_element(self->file, "x:TextVAlign", "Center", NULL);
/// }

/*
 * Write the <x:FmlaMacro> element.
 */
/// STATIC void
/// _vml_write_fmla_macro(lxw_vml *self, lxw_vml_obj *vml_obj)
/// {
///     lxw_xml_data_element(self->file, "x:FmlaMacro", vml_obj->macro, NULL);
/// }

/*
 * Write the <x:PrintObject> element.
 */
/// STATIC void
/// _vml_write_print_object(lxw_vml *self)
/// {
///     lxw_xml_data_element(self->file, "x:PrintObject", "False", NULL);
/// }

std::string vml_t::write_aspect_ratio_lock() const
{
  return xml_empty_tag("o:lock", {
                                     {"v:ext",       "edit"},
                                     {"aspectratio", "t"   },
  });
}

std::string vml_t::write_rotation_lock() const
{
  return xml_empty_tag("o:lock", {
                                     {"v:ext",    "edit"},
                                     {"rotation", "t"   },
  });
}

std::string vml_t::write_column(const vml_obj_t& vml_obj) const
{
  return xml_data_element("x:Column", std::to_string(vml_obj.col_));
}

std::string vml_t::write_row(const vml_obj_t& vml_obj) const
{
  return xml_data_element("x:Row", std::to_string(vml_obj.row_));
}

std::string vml_t::write_auto_fill() const
{
  return xml_data_element("x:AutoFill", "False");
}

std::string vml_t::write_anchor(const vml_obj_t& vml_obj) const
{
  return xml_data_element("x:Anchor", std::format("{}, {}, {}, {}, {}, {}, {}, {}", vml_obj.from_.col_,
                                                  static_cast<uint32_t>(vml_obj.from_.col_offset_), vml_obj.from_.row_,
                                                  static_cast<uint32_t>(vml_obj.from_.row_offset_), vml_obj.to_.col_,
                                                  static_cast<uint32_t>(vml_obj.to_.col_offset_), vml_obj.to_.row_,
                                                  static_cast<uint32_t>(vml_obj.to_.row_offset_)));
}

std::string vml_t::write_size_with_cells() const
{
  return xml_empty_tag("x:SizeWithCells");
}

std::string vml_t::write_move_with_cells() const
{
  return xml_empty_tag("x:MoveWithCells");
}

std::string vml_t::write_shadow() const
{
  return xml_empty_tag("v:shadow", {
                                       {"on",       "t"    },
                                       {"color",    "black"},
                                       {"obscured", "t"    },
  });
}

std::string vml_t::write_stroke() const
{
  return xml_empty_tag("v:stroke", {
                                       {"joinstyle", "miter"}
  });
}

/*
 * Write the <o:lock> element.
 */
/// STATIC void
/// _vml_write_shapetype_lock(lxw_vml *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("v:ext", "edit");
///     LXW_PUSH_ATTRIBUTES_STR("shapetype", "t");
///
///     lxw_xml_empty_tag(self->file, "o:lock", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _vml_write_font(lxw_vml *self, lxw_vml_obj *vml_obj)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("face", "Calibri");
///     LXW_PUSH_ATTRIBUTES_STR("size", "220");
///     LXW_PUSH_ATTRIBUTES_STR("color", "#000000");
///
///     lxw_xml_data_element(self->file, "font", vml_obj->name, &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string vml_t::write_imagedata(uint32_t rel_index, const std::string& name) const
{
  return xml_empty_tag("v:imagedata", {
                                          {"o:relid", std::format("rId{}", rel_index)},
                                          {"o:title", name},
  });
}

std::string vml_t::write_image_path() const
{
  return xml_empty_tag("v:path", {
                                     {"o:extrusionok",   "f"   },
                                     {"gradientshapeok", "t"   },
                                     {"o:connecttype",   "rect"},
  });
}

std::string vml_t::write_image_shape(uint32_t vml_shape_id, uint32_t z_index, const vml_obj_t& image_obj) const
{
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     char width_str[LXW_ATTR_32];
  ///     char height_str[LXW_ATTR_32];
  ///     char style[LXW_MAX_ATTRIBUTE_LENGTH];
  ///     char o_spid[LXW_ATTR_32];
  ///     char type[] = ;

  // Scale the height/width by the resolution, relative to 72dpi.
  double width  = image_obj.width_ * (72.0 / image_obj.x_dpi_);
  double height = image_obj.height_ * (72.0 / image_obj.y_dpi_);

  // Excel uses a rounding based around 72 and 96 dpi.
  width  = 72.0 / 96.0 * static_cast<uint32_t>(width * 96.0 / 72 + 0.25);
  height = 72.0 / 96.0 * static_cast<uint32_t>(height * 96.0 / 72 + 0.25);

  std::string xml_data = xml_start_tag(
      "v:shape",
      {
          {"id", image_obj.image_position_},
          {"o:spid", std::format("_x0000_s{}", vml_shape_id)},
          {"type", "#_x0000_t75"},
          {"style", std::format("position:absolute;margin-left:0;margin-top:0;width:{}pt;height:{}pt;z-index:{}", width,
           height, z_index)},
  });
  xml_data += write_imagedata(image_obj.rel_index_, image_obj.name_);
  xml_data += write_rotation_lock();
  xml_data += xml_end_tag("v:shape");

  return xml_data;
}

std::string vml_t::write_image_shapetype() const
{
  std::string xml_data = xml_start_tag("v:shapetype", {
                                                          {"id",               "_x0000_t75"            },
                                                          {"coordsize",        "21600,21600"           },
                                                          {"o:spt",            "75"                    },
                                                          {"o:preferrelative", "t"                     },
                                                          {"path",             "m@4@5l@4@11@9@11@9@5xe"},
                                                          {"filled",           "f"                     },
                                                          {"stroked",          "f"                     },
  });
  xml_data += write_stroke();
  xml_data += write_formulas();
  xml_data += write_image_path();
  xml_data += write_aspect_ratio_lock();
  xml_data += xml_end_tag("v:shapetype");

  return xml_data;
}

/*
 * Write the <x:ClientData> element.
 */
/// STATIC void
/// _vml_write_button_client_data(lxw_vml *self, lxw_vml_obj *vml_obj)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("ObjectType", "Button");
///
///     lxw_xml_start_tag(self->file, "x:ClientData", &attributes);
///
///     /* Write the <x:Anchor> element. */
///     _vml_write_anchor(self, vml_obj);
///
///     /* Write the x:PrintObject element. */
///     _vml_write_print_object(self);
///
///     /* Write the x:AutoFill element. */
///     _vml_write_auto_fill(self);
///
///     /* Write the x:FmlaMacro element. */
///     _vml_write_fmla_macro(self, vml_obj);
///
///     /* Write the x:TextHAlign element. */
///     _vml_write_text_halign(self);
///
///     /* Write the x:TextVAlign element. */
///     _vml_write_text_valign(self);
///
///     lxw_xml_end_tag(self->file, "x:ClientData");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <div> element.
 */
/// STATIC void
/// _vml_write_button_div(lxw_vml *self, lxw_vml_obj *vml_obj)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("style", "text-align:center");
///
///     lxw_xml_start_tag(self->file, "div", &attributes);
///
///     /* Write the font element. */
///     _vml_write_font(self, vml_obj);
///
///     lxw_xml_end_tag(self->file, "div");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <v:textbox> element.
 */
/// STATIC void
/// _vml_write_button_textbox(lxw_vml *self, lxw_vml_obj *vml_obj)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("style", "mso-direction-alt:auto");
///     LXW_PUSH_ATTRIBUTES_STR("o:singleclick", "f");
///
///     lxw_xml_start_tag(self->file, "v:textbox", &attributes);
///
///     /* Write the div element. */
///     _vml_write_button_div(self, vml_obj);
///
///     lxw_xml_end_tag(self->file, "v:textbox");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <v:fill> element.
 */
/// STATIC void
/// _vml_write_button_fill(lxw_vml *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("color2", "buttonFace [67]");
///     LXW_PUSH_ATTRIBUTES_STR("o:detectmouseclick", "t");
///
///     lxw_xml_empty_tag(self->file, "v:fill", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <v:path> element for buttons.
 */
/// STATIC void
/// _vml_write_button_path(lxw_vml *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("shadowok", "f");
///     LXW_PUSH_ATTRIBUTES_STR("o:extrusionok", "f");
///     LXW_PUSH_ATTRIBUTES_STR("strokeok", "f");
///     LXW_PUSH_ATTRIBUTES_STR("fillok", "f");
///     LXW_PUSH_ATTRIBUTES_STR("o:connecttype", "rect");
///
///     lxw_xml_empty_tag(self->file, "v:path", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <v:shape> element for buttons.
 */
/// STATIC void
/// _vml_write_button_shape(lxw_vml *self, uint32_t vml_shape_id,
///                         uint32_t z_index, lxw_vml_obj *vml_obj)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char type[] = "#_x0000_t201";
///     char o_button[] = "t";
///     char fillcolor[] = "buttonFace [67]";
///     char strokecolor[] = "windowText [64]";
///     char o_insetmode[] = "auto";
///
///     char id[LXW_ATTR_32];
///     char margin_left[LXW_ATTR_32];
///     char margin_top[LXW_ATTR_32];
///     char width[LXW_ATTR_32];
///     char height[LXW_ATTR_32];
///     char style[LXW_MAX_ATTRIBUTE_LENGTH];
///
///     lxw_sprintf_dbl(margin_left, vml_obj->col_absolute * 0.75);
///     lxw_sprintf_dbl(margin_top, vml_obj->row_absolute * 0.75);
///     lxw_sprintf_dbl(width, vml_obj->width * 0.75);
///     lxw_sprintf_dbl(height, vml_obj->height * 0.75);
///
///     lxw_snprintf(id, LXW_ATTR_32, "_x0000_s%d", vml_shape_id);
///
///     lxw_snprintf(style,
///                  LXW_MAX_ATTRIBUTE_LENGTH,
///                  "position:absolute;"
///                  "margin-left:%spt;"
///                  "margin-top:%spt;"
///                  "width:%spt;"
///                  "height:%spt;"
///                  "z-index:%d;"
///                  "mso-wrap-style:tight",
///                  margin_left, margin_top, width, height, z_index);
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("id", id);
///     LXW_PUSH_ATTRIBUTES_STR("type", type);
///
///     if (vml_obj->text)
///         LXW_PUSH_ATTRIBUTES_STR("alt", vml_obj->text);
///
///     LXW_PUSH_ATTRIBUTES_STR("style", style);
///     LXW_PUSH_ATTRIBUTES_STR("o:button", o_button);
///     LXW_PUSH_ATTRIBUTES_STR("fillcolor", fillcolor);
///     LXW_PUSH_ATTRIBUTES_STR("strokecolor", strokecolor);
///     LXW_PUSH_ATTRIBUTES_STR("o:insetmode", o_insetmode);
///
///     lxw_xml_start_tag(self->file, "v:shape", &attributes);
///
///     /* Write the v:fill element. */
///     _vml_write_button_fill(self);
///
///     /* Write the o:lock element. */
///     _vml_write_rotation_lock(self);
///
///     /* Write the v:textbox element. */
///     _vml_write_button_textbox(self, vml_obj);
///
///     /* Write the x:ClientData element. */
///     _vml_write_button_client_data(self, vml_obj);
///
///     lxw_xml_end_tag(self->file, "v:shape");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write the <v:shapetype> element for buttons.
 */
/// STATIC void
/// _vml_write_button_shapetype(lxw_vml *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char id[] = "_x0000_t201";
///     char coordsize[] = "21600,21600";
///     char o_spt[] = "201";
///     char path[] = "m,l,21600r21600,l21600,xe";
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("id", id);
///     LXW_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
///     LXW_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
///     LXW_PUSH_ATTRIBUTES_STR("path", path);
///
///     lxw_xml_start_tag(self->file, "v:shapetype", &attributes);
///
///     /* Write the v:stroke element. */
///     _vml_write_stroke(self);
///
///     /* Write the v:path element. */
///     _vml_write_button_path(self);
///
///     /* Write the o:lock element. */
///     _vml_write_shapetype_lock(self);
///
///     lxw_xml_end_tag(self->file, "v:shapetype");
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string vml_t::write_comment_client_data(const vml_obj_t& vml_obj) const
{
  std::string xml_data = xml_start_tag("x:ClientData", {
                                                           {"ObjectType", "Note"}
  });
  xml_data += write_move_with_cells();
  xml_data += write_size_with_cells();
  xml_data += write_anchor(vml_obj);
  xml_data += write_auto_fill();
  xml_data += write_row(vml_obj);
  xml_data += write_column(vml_obj);
  if(vml_obj.visible_ == comment_display_t::VISIBLE)
  {
    xml_data += write_visible();
  }
  xml_data += xml_end_tag("x:ClientData");

  return xml_data;
}

std::string vml_t::write_comment_div() const
{
  std::string xml_data = xml_start_tag("div", {
                                                  {"style", "text-align:left"}
  });
  xml_data += xml_end_tag("div");
  return xml_data;
}

std::string vml_t::write_comment_textbox() const
{
  std::string xml_data = xml_start_tag("v:textbox", {
                                                        {"style", "mso-direction-alt:auto"}
  });
  xml_data += write_comment_div();
  xml_data += xml_end_tag("v:textbox");

  return xml_data;
}

std::string vml_t::write_comment_fill() const
{
  return xml_empty_tag("v:fill", {
                                     {"color2", "#ffffe1"}
  });
}

std::string vml_t::write_comment_path(bool has_gradient, const std::string& type) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(has_gradient)
  {
    attributes.emplace_back("gradientshapeok", "t");
  }
  attributes.emplace_back("o:connecttype", type);

  return xml_empty_tag("v:path", attributes);
}

std::string vml_t::write_comment_shape(uint32_t vml_shape_id, uint32_t z_index, vml_obj_t& vml_obj) const
{
  if(vml_obj.visible_ == comment_display_t::DEFAULT)
  {
    vml_obj.visible_ = comment_display_default_;
  }

  const std::string visible   = vml_obj.visible_ == comment_display_t::VISIBLE ? "visible" : "hidden";
  const std::string fillcolor = vml_obj.color_ != color_t::UNSET
                                    ? std::format("#{:06x}", static_cast<uint32_t>(vml_obj.color_) & COLOR_MASK)
                                    : "#ffffe1";

  std::string xml_data = xml_start_tag(
      "v:shape", {
                     {"id", std::format("_x0000_s{}", vml_shape_id)},
                     {"type", "#_x0000_t202"},
                     {"style", std::format("position:absolute;"
                                           "margin-left:{}pt;"
                                           "margin-top:{}pt;"
                                           "width:{}pt;"
                                           "height:{}pt;"
                                           "z-index:{};"
                                           "visibility:{}", vml_obj.col_absolute_ * 0.75, vml_obj.row_absolute_ * 0.75,
                      vml_obj.width_ * 0.75, vml_obj.height_ * 0.75, z_index, visible)},
                     {"fillcolor", fillcolor},
                     {"o:insetmode", "auto"},
  });
  xml_data += write_comment_fill();
  xml_data += write_shadow();
  xml_data += write_comment_path(false, "none");
  xml_data += write_comment_textbox();
  xml_data += write_comment_client_data(vml_obj);
  xml_data += xml_end_tag("v:shape");

  return xml_data;
}

std::string vml_t::write_comment_shapetype() const
{
  std::string xml_data = xml_start_tag("v:shapetype", {
                                                          {"id",        "_x0000_t202"              },
                                                          {"coordsize", "21600,21600"              },
                                                          {"o:spt",     "202"                      },
                                                          {"path",      "m,l,21600r21600,l21600,xe"},
  });
  xml_data += write_stroke();
  xml_data += write_comment_path(true, "rect");
  xml_data += xml_end_tag("v:shapetype");

  return xml_data;
}

std::string vml_t::write_idmap() const
{
  return std::format("<o:idmap v:ext=\"edit\" data=\"{}\"/>", vml_data_id_str_);
}

std::string vml_t::write_shapelayout() const
{
  std::string xml_data = xml_start_tag("o:shapelayout", {
                                                            {"v:ext", "edit"}
  });
  xml_data += write_idmap();
  xml_data += xml_end_tag("o:shapelayout");

  return xml_data;
}

std::string vml_t::write_xml_namespace() const
{
  return xml_start_tag("xml", {
                                  {"xmlns:v", "urn:schemas-microsoft-com:vml"          },
                                  {"xmlns:o", "urn:schemas-microsoft-com:office:office"},
                                  {"xmlns:x", "urn:schemas-microsoft-com:office:excel" },
  });
}

std::string vml_t::assemble_xml_file()
{
  ///     lxw_vml_obj *comment_obj;
  ///     lxw_vml_obj *button_obj;
  ///     lxw_vml_obj *image_obj;
  uint32_t z_index = 1;

  std::string xml_data = write_xml_namespace();
  xml_data += write_shapelayout();
  ///
  ///     if (self->button_objs && !STAILQ_EMPTY(self->button_objs)) {
  ///         /* Write the <v:shapetype> element. */
  ///         _vml_write_button_shapetype(self);
  ///
  ///         STAILQ_FOREACH(button_obj, self->button_objs, list_pointers) {
  ///             self->vml_shape_id++;
  ///
  ///             /* Write the <v:shape> element. */
  ///             _vml_write_button_shape(self, self->vml_shape_id, z_index,
  ///                                     button_obj);
  ///
  ///             z_index++;
  ///         }
  ///     }

  if(!comment_objs_.empty())
  {
    xml_data += write_comment_shapetype();
    for(auto& comment: comment_objs_)
    {
      vml_shape_id_++;
      xml_data += write_comment_shape(vml_shape_id_, z_index, comment);
      z_index++;
    }
  }

  if(!image_objs_.empty())
  {
    xml_data += write_image_shapetype();
    for(const auto& image_obj: image_objs_)
    {
      vml_shape_id_++;
      xml_data += write_image_shape(vml_shape_id_, z_index, image_obj);

      z_index++;
    }
  }
  xml_data += xml_end_tag("xml");

  return xml_data;
}

}
