/*
 * drawing - A library for creating Excel XLSX drawing files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/drawing.h"

#include "xwpp/common.h"
#include "xwpp/utility.h"
#include "xwpp/worksheet.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <optional>

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_drawing_workspace()
{
  return xml_start_tag("xdr:wsDr", {
                                     {"xmlns:xdr", SCHEMA_DRAWING + "/spreadsheetDrawing"},
                                     {"xmlns:a",   SCHEMA_DRAWING + "/main"              },
  });
}

[[nodiscard]] std::string write_col(const std::string& data)
{
  return xml_data_element("xdr:col", data);
}

[[nodiscard]] std::string write_col_off(const std::string& data)
{
  return xml_data_element("xdr:colOff", data);
}

[[nodiscard]] std::string write_row(const std::string& data)
{
  return xml_data_element("xdr:row", data);
}

[[nodiscard]] std::string write_row_off(const std::string& data)
{
  return xml_data_element("xdr:rowOff", data);
}

[[nodiscard]] std::string write_uri_ext(const std::string& uri)
{
  return xml_start_tag("a:ext", {
                                  {"uri", uri}
  });
}

[[nodiscard]] std::string write_a16_creation_id()
{
  return xml_empty_tag("a16:creationId", {
                                           {"xmlns:a16", "http://schemas.microsoft.com/office/drawing/2014/main"},
                                           {"id",        "{00000000-0008-0000-0000-000002000000}"               },
  });
}

[[nodiscard]] std::string write_adec_decorative()
{
  return xml_empty_tag("adec:decorative",
                       {
                         {"xmlns:adec", "http://schemas.microsoft.com/office/drawing/2017/decorative"},
                         {"val",        "1"                                                          },
  });
}

[[nodiscard]] std::string write_a_hlink_click(uint32_t rel_index, const std::string& tip)
{
  std::vector<std::tuple<std::string, std::string>> attributes{
    {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
    {"r:id", std::format("rId{}", rel_index)},
  };

  if(!tip.empty())
  {
    attributes.emplace_back("tooltip", tip);
  }

  return xml_empty_tag("a:hlinkClick", attributes);
}

[[nodiscard]] std::string write_a_pic_locks()
{
  return xml_empty_tag("a:picLocks", {
                                       {"noChangeAspect", "1"}
  });
}

[[nodiscard]] std::string write_a_blip(uint32_t index)
{
  return xml_empty_tag("a:blip", {
                                   {"xmlns:r", SCHEMA_OFFICEDOC + "/relationships"},
                                   {"r:embed", std::format("rId{}", index)},
  });
}

[[nodiscard]] std::string write_a_fill_rect()
{
  return xml_empty_tag("a:fillRect");
}

[[nodiscard]] std::string write_a_ext(const drawing_object_t& drawing_object)
{
  return xml_empty_tag("a:ext", {
                                  {"cx", std::to_string(drawing_object.width_) },
                                  {"cy", std::to_string(drawing_object.height_)},
  });
}

[[nodiscard]] std::string write_a_off(const drawing_object_t& drawing_object)
{
  return xml_empty_tag("a:off", {
                                  {"x", std::to_string(drawing_object.col_absolute_)},
                                  {"y", std::to_string(drawing_object.row_absolute_)},
  });
}

[[nodiscard]] std::string write_a_av_lst()
{
  return xml_empty_tag("a:avLst");
}

[[nodiscard]] std::string write_client_data()
{
  return xml_empty_tag("xdr:clientData");
}

[[nodiscard]] std::string write_a_graphic_frame_locks()
{
  return xml_empty_tag("a:graphicFrameLocks", {
                                                {"noGrp", "1"}
  });
}

[[nodiscard]] std::string write_xfrm_offset()
{
  return xml_empty_tag("a:off", {
                                  {"x", "0"},
                                  {"y", "0"},
  });
}

[[nodiscard]] std::string write_xfrm_extension()
{
  return xml_empty_tag("a:ext", {
                                  {"cx", "0"},
                                  {"cy", "0"},
  });
}

[[nodiscard]] std::string write_xfrm()
{
  std::string xml_data = xml_start_tag("xdr:xfrm");
  xml_data += write_xfrm_offset();
  xml_data += write_xfrm_extension();
  xml_data += xml_end_tag("xdr:xfrm");

  return xml_data;
}

[[nodiscard]] std::string write_chart(uint32_t index)
{
  return xml_empty_tag("c:chart", {
                                    {"xmlns:c", SCHEMA_DRAWING + "/chart"},
                                    {"xmlns:r", SCHEMA_OFFICEDOC + "/relationships"},
                                    {"r:id", std::format("rId{}", index)},
  });
}

[[nodiscard]] std::string write_ext(uint32_t cx, uint32_t cy)
{
  return xml_empty_tag("xdr:ext", {
                                    {"cx", std::to_string(cx)},
                                    {"cy", std::to_string(cy)},
  });
}

// NOLINTNEXTLINE(readability-identifier-length)
[[nodiscard]] std::string write_pos(int32_t x, int32_t y)
{
  return xml_empty_tag("xdr:pos", {
                                    {"x", std::to_string(x)},
                                    {"y", std::to_string(y)},
  });
}

[[nodiscard]] std::string write_coords(const drawing_coords_t& coords)
{
  std::string xml_data = write_col(std::to_string(coords.col_));
  xml_data += write_col_off(std::to_string(static_cast<uint32_t>(coords.col_offset_)));
  xml_data += write_row(std::to_string(coords.row_));
  xml_data += write_row_off(std::to_string(static_cast<uint32_t>(coords.row_offset_)));

  return xml_data;
}

[[nodiscard]] std::string write_from(const drawing_coords_t& coords)
{
  std::string xml_data = xml_start_tag("xdr:from");
  xml_data += write_coords(coords);
  xml_data += xml_end_tag("xdr:from");

  return xml_data;
}

[[nodiscard]] std::string write_to(const drawing_coords_t& coords)
{
  std::string xml_data = xml_start_tag("xdr:to");
  xml_data += write_coords(coords);
  xml_data += xml_end_tag("xdr:to");

  return xml_data;
}

[[nodiscard]] std::string write_c_nv_pic_pr()
{
  std::string xml_data = xml_start_tag("xdr:cNvPicPr");
  xml_data += write_a_pic_locks();
  xml_data += xml_end_tag("xdr:cNvPicPr");

  return xml_data;
}

[[nodiscard]] std::string write_decorative()
{
  std::string xml_data = xml_start_tag("a:extLst");

  xml_data += write_uri_ext("{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");
  xml_data += write_a16_creation_id();
  xml_data += xml_end_tag("a:ext");

  xml_data += write_uri_ext("{C183D7F6-B498-43B3-948B-1728B52AA6E4}");
  xml_data += write_adec_decorative();
  xml_data += xml_end_tag("a:ext");

  xml_data += xml_end_tag("a:extLst");

  return xml_data;
}

[[nodiscard]] std::string write_c_nv_pr(const std::string& object_name, uint32_t index,
                                        const std::optional<drawing_object_t>& drawing_object)
{
  std::vector<std::tuple<std::string, std::string>> attributes{
    {"id", std::to_string(index + 1)},
    {"name", std::format("{} {}", object_name, index)},
  };

  if(drawing_object && !drawing_object->description_.empty() && !drawing_object->decorative_)
  {
    attributes.emplace_back("descr", drawing_object->description_);
  }

  if(drawing_object && (drawing_object->url_rel_index_ != 0 || drawing_object->decorative_))
  {
    std::string xml_data = xml_start_tag("xdr:cNvPr", attributes);

    if(drawing_object->url_rel_index_ != 0)
    {
      xml_data += write_a_hlink_click(drawing_object->url_rel_index_, drawing_object->tip_);
    }

    if(drawing_object->decorative_)
    {
      xml_data += write_decorative();
    }
    xml_data += xml_end_tag("xdr:cNvPr");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("xdr:cNvPr", attributes);
  }
}

[[nodiscard]] std::string write_nv_pic_pr(uint32_t index, const drawing_object_t& drawing_object)
{
  std::string xml_data = xml_start_tag("xdr:nvPicPr");
  xml_data += write_c_nv_pr("Picture", index, drawing_object);
  xml_data += write_c_nv_pic_pr();
  xml_data += xml_end_tag("xdr:nvPicPr");

  return xml_data;
}

[[nodiscard]] std::string write_a_prst_geom()
{
  std::string xml_data = xml_start_tag("a:prstGeom", {
                                                       {"prst", "rect"}
  });
  xml_data += write_a_av_lst();
  xml_data += xml_end_tag("a:prstGeom");

  return xml_data;
}

[[nodiscard]] std::string write_a_xfrm(const drawing_object_t& drawing_object)
{
  std::string xml_data = xml_start_tag("a:xfrm");
  xml_data += write_a_off(drawing_object);
  xml_data += write_a_ext(drawing_object);
  xml_data += xml_end_tag("a:xfrm");

  return xml_data;
}

[[nodiscard]] std::string write_sp_pr(const drawing_object_t& drawing_object)
{
  std::string xml_data = xml_start_tag("xdr:spPr");
  xml_data += write_a_xfrm(drawing_object);
  xml_data += write_a_prst_geom();
  xml_data += xml_end_tag("xdr:spPr");

  return xml_data;
}

[[nodiscard]] std::string write_a_stretch()
{
  std::string xml_data = xml_start_tag("a:stretch");
  xml_data += write_a_fill_rect();
  xml_data += xml_end_tag("a:stretch");

  return xml_data;
}

[[nodiscard]] std::string write_blip_fill(uint32_t index)
{
  std::string xml_data = xml_start_tag("xdr:blipFill");
  xml_data += write_a_blip(index);
  xml_data += write_a_stretch();
  xml_data += xml_end_tag("xdr:blipFill");

  return xml_data;
}

[[nodiscard]] std::string write_pic(uint32_t index, const drawing_object_t& drawing_object)
{
  std::string xml_data = xml_start_tag("xdr:pic");
  xml_data += write_nv_pic_pr(index, drawing_object);
  xml_data += write_blip_fill(drawing_object.rel_index_);
  xml_data += write_sp_pr(drawing_object);
  xml_data += xml_end_tag("xdr:pic");

  return xml_data;
}

[[nodiscard]] std::string write_a_graphic_data(uint32_t index)
{
  std::string xml_data = xml_start_tag("a:graphicData", {
                                                          {"uri", SCHEMA_DRAWING + "/chart"}
  });
  xml_data += write_chart(index);
  xml_data += xml_end_tag("a:graphicData");

  return xml_data;
}

[[nodiscard]] std::string write_a_graphic(uint32_t index)
{
  std::string xml_data = xml_start_tag("a:graphic");
  xml_data += write_a_graphic_data(index);
  xml_data += xml_end_tag("a:graphic");

  return xml_data;
}

}

void drawing_t::add_drawing_object(const drawing_object_t& drawing_object)
{
  drawing_objects_.push_back(drawing_object);
}

std::string drawing_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_drawing_workspace();

  if(embedded_)
  {
    uint32_t index = 1;

    for(const auto& drawing_object: drawing_objects_)
    {
      xml_data += write_two_cell_anchor(index, drawing_object);
      index++;
    }
  }
  else
  {
    xml_data += write_absolute_anchor(1);
  }
  xml_data += xml_end_tag("xdr:wsDr");

  return xml_data;
}

std::string drawing_t::write_absolute_anchor(uint32_t frame_index) const
{
  std::string xml_data = xml_start_tag("xdr:absoluteAnchor");

  if(orientation_ == drawing_orientation_t::LANDSCAPE)
  {
    xml_data += write_pos(0, 0);
    xml_data += write_ext(9308969, 6078325);
  }
  else
  {
    xml_data += write_pos(0, -47625);
    xml_data += write_ext(6162675, 6124575);
  }
  xml_data += write_graphic_frame(frame_index, frame_index, std::nullopt);
  xml_data += write_client_data();
  xml_data += xml_end_tag("xdr:absoluteAnchor");

  return xml_data;
}

std::string drawing_t::write_two_cell_anchor(uint32_t index, const drawing_object_t& drawing_object) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  if(drawing_object.anchor_ == static_cast<uint8_t>(object_position_t::MOVE_DONT_SIZE))
  {
    attributes.emplace_back("editAs", "oneCell");
  }
  else if(drawing_object.anchor_ == static_cast<uint8_t>(object_position_t::DONT_MOVE_DONT_SIZE))
  {
    attributes.emplace_back("editAs", "absolute");
  }

  std::string xml_data = xml_start_tag("xdr:twoCellAnchor", attributes);

  xml_data += write_from(drawing_object.from_);
  xml_data += write_to(drawing_object.to_);

  if(drawing_object.type_ == drawing_types_t::CHART)
  {
    xml_data += write_graphic_frame(index, drawing_object.rel_index_, drawing_object);
  }
  else if(drawing_object.type_ == drawing_types_t::IMAGE)
  {
    xml_data += write_pic(index, drawing_object);
  }
  else
  {
    // TODO ???
    /* Write the xdr:sp element for shapes. */
    /* _drawing_write_sp(self, index, col_absolute, row_absolute, width, height,  shape); */
  }
  xml_data += write_client_data();
  xml_data += xml_end_tag("xdr:twoCellAnchor");

  return xml_data;
}

std::string drawing_t::write_c_nv_graphic_frame_pr() const
{
  if(embedded_)
  {
    return xml_empty_tag("xdr:cNvGraphicFramePr");
  }
  else
  {
    std::string xml_data = xml_start_tag("xdr:cNvGraphicFramePr");
    xml_data += write_a_graphic_frame_locks();
    xml_data += xml_end_tag("xdr:cNvGraphicFramePr");

    return xml_data;
  }
}

std::string drawing_t::write_nv_graphic_frame_pr(uint32_t index,
                                                 const std::optional<drawing_object_t>& drawing_object) const
{
  std::string xml_data = xml_start_tag("xdr:nvGraphicFramePr");
  xml_data += write_c_nv_pr("Chart", index, drawing_object);
  xml_data += write_c_nv_graphic_frame_pr();
  xml_data += xml_end_tag("xdr:nvGraphicFramePr");

  return xml_data;
}

std::string drawing_t::write_graphic_frame(uint32_t index, uint32_t rel_index,
                                           const std::optional<drawing_object_t>& drawing_object) const
{
  std::string xml_data = xml_start_tag("xdr:graphicFrame", {
                                                             {"macro", ""}
  });
  xml_data += write_nv_graphic_frame_pr(index, drawing_object);
  xml_data += write_xfrm();
  xml_data += write_a_graphic(rel_index);
  xml_data += xml_end_tag("xdr:graphicFrame");

  return xml_data;
}

}
