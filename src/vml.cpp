/*
 * vml - A library for creating Excel XLSX vml files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
vml_t::vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& comment_objs,
             const std::vector<vml_obj_t>& button_objs, uint32_t vml_shape_id,
             comment_display_t comment_display_default)
  : button_objs_{button_objs}
  , comment_objs_{comment_objs}
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

std::string vml_t::write_text_halign() const
{
  return xml_data_element("x:TextHAlign", "Center");
}

std::string vml_t::write_text_valign() const
{
  return xml_data_element("x:TextVAlign", "Center");
}

std::string vml_t::write_fmla_macro(const vml_obj_t& vml_obj) const
{
  return xml_data_element("x:FmlaMacro", vml_obj.macro_);
}

std::string vml_t::write_print_object() const
{
  return xml_data_element("x:PrintObject", "False");
}

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

std::string vml_t::write_shapetype_lock() const
{
  return xml_empty_tag("o:lock", {
                                     {"v:ext",     "edit"},
                                     {"shapetype", "t"   },
  });
}

std::string vml_t::write_font(const vml_obj_t& vml_obj) const
{
  return xml_data_element("font", vml_obj.name_,
                          {
                              {"face",  "Calibri"},
                              {"size",  "220"    },
                              {"color", "#000000"},
  });
}

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

std::string vml_t::write_button_client_data(const vml_obj_t& vml_obj) const
{
  std::string xml_data = xml_start_tag("x:ClientData", {
                                                           {"ObjectType", "Button"}
  });
  xml_data += write_anchor(vml_obj);
  xml_data += write_print_object();
  xml_data += write_auto_fill();
  xml_data += write_fmla_macro(vml_obj);
  xml_data += write_text_halign();
  xml_data += write_text_valign();
  xml_data += xml_end_tag("x:ClientData");

  return xml_data;
}

std::string vml_t::write_button_div(const vml_obj_t& vml_obj) const
{
  std::string xml_data = xml_start_tag("div", {
                                                  {"style", "text-align:center"}
  });
  xml_data += write_font(vml_obj);
  xml_data += xml_end_tag("div");

  return xml_data;
}

std::string vml_t::write_button_textbox(const vml_obj_t& vml_obj) const
{
  std::string xml_data = xml_start_tag("v:textbox", {
                                                        {"style",         "mso-direction-alt:auto"},
                                                        {"o:singleclick", "f"                     },
  });
  xml_data += write_button_div(vml_obj);
  xml_data += xml_end_tag("v:textbox");

  return xml_data;
}

std::string vml_t::write_button_fill() const
{
  return xml_empty_tag("v:fill", {
                                     {"color2",             "buttonFace [67]"},
                                     {"o:detectmouseclick", "t"              },
  });
}

std::string vml_t::write_button_path() const
{
  return xml_empty_tag("v:path", {
                                     {"shadowok",      "f"   },
                                     {"o:extrusionok", "f"   },
                                     {"strokeok",      "f"   },
                                     {"fillok",        "f"   },
                                     {"o:connecttype", "rect"},
  });
}

std::string vml_t::write_button_shape(uint32_t vml_shape_id, uint32_t z_index, const vml_obj_t& vml_obj) const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"id", std::format("_x0000_s{}", vml_shape_id)},
      {"type", "#_x0000_t201"},
  };

  if(!vml_obj.text_.empty())
  {
    attributes.emplace_back("alt", vml_obj.text_);
  }
  attributes.emplace_back(
      "style",
      std::format(
          "position:absolute;margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{};mso-wrap-style:tight",
          vml_obj.col_absolute_ * 0.75, vml_obj.row_absolute_ * 0.75, vml_obj.width_ * 0.75, vml_obj.height_ * 0.75,
          z_index));
  attributes.emplace_back("o:button", "t");
  attributes.emplace_back("fillcolor", "buttonFace [67]");
  attributes.emplace_back("strokecolor", "windowText [64]");
  attributes.emplace_back("o:insetmode", "auto");

  std::string xml_data = xml_start_tag("v:shape", attributes);
  xml_data += write_button_fill();
  xml_data += write_rotation_lock();
  xml_data += write_button_textbox(vml_obj);
  xml_data += write_button_client_data(vml_obj);
  xml_data += xml_end_tag("v:shape");

  return xml_data;
}

std::string vml_t::write_button_shapetype() const
{
  std::string xml_data = xml_start_tag("v:shapetype", {
                                                          {"id",        "_x0000_t201"              },
                                                          {"coordsize", "21600,21600"              },
                                                          {"o:spt",     "201"                      },
                                                          {"path",      "m,l,21600r21600,l21600,xe"},
  });
  xml_data += write_stroke();
  xml_data += write_button_path();
  xml_data += write_shapetype_lock();
  xml_data += xml_end_tag("v:shapetype");

  return xml_data;
}

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
  uint32_t z_index = 1;

  std::string xml_data = write_xml_namespace();
  xml_data += write_shapelayout();

  if(!button_objs_.empty())
  {
    xml_data += write_button_shapetype();
    for(const auto& button_obj: button_objs_)
    {
      vml_shape_id_++;
      xml_data += write_button_shape(vml_shape_id_, z_index, button_obj);
      z_index++;
    }
  }

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
