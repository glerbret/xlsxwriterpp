/*
 * drawing - A Xlsxwriter++ library for creating Excel XLSX drawing files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_DRAWING_H
#define XWPP_DRAWING_H

#include "xwpp/common.h"

#include <cstdint>

namespace xwpp
{

enum class drawing_types_t
{
  NONE,
  IMAGE,
  CHART,
  SHAPE
};

enum class image_types_t
{
  UNKNOWN,
  PNG,
  JPEG,
  BMP,
  GIF
};

// Coordinates used in a drawing object.
struct drawing_coords_t
{
  uint32_t col_;
  uint32_t row_;
  double col_offset_;
  double row_offset_;
};

// Object to represent the properties of a drawing.
struct drawing_object_t
{
  drawing_types_t type_ = drawing_types_t::NONE;
  uint8_t anchor_       = 0; // TODO Use object_position_t (but crossed inclusion for the moment)
  drawing_coords_t from_;
  drawing_coords_t to_;
  uint64_t col_absolute_;
  uint64_t row_absolute_;
  uint32_t width_;
  uint32_t height_;
  // TODO ?     uint8_t shape;
  uint32_t rel_index_;
  uint32_t url_rel_index_;
  std::string description_;
  std::string tip_;
  bool decorative_;
};

class drawing_t
{
public:
  void add_drawing_object(const drawing_object_t& drawing_object);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  friend class worksheet_t;

  [[nodiscard]] std::string write_drawing_workspace() const;
  [[nodiscard]] std::string write_absolute_anchor(uint32_t frame_index) const;
  [[nodiscard]] std::string write_col(const std::string& data) const;
  [[nodiscard]] std::string write_col_off(const std::string& data) const;
  [[nodiscard]] std::string write_row(const std::string& data) const;
  [[nodiscard]] std::string write_row_off(const std::string& data) const;
  [[nodiscard]] std::string write_coords(const drawing_coords_t& coords) const;
  [[nodiscard]] std::string write_from(const drawing_coords_t& coords) const;
  [[nodiscard]] std::string write_to(const drawing_coords_t& coords) const;
  [[nodiscard]] std::string write_two_cell_anchor(uint32_t index, const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_pic(uint32_t index, const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_nv_pic_pr(uint32_t index, const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_c_nv_pr(const std::string& object_name, uint32_t index,
                                          const std::optional<drawing_object_t>& drawing_object) const;
  [[nodiscard]] std::string write_decorative() const;
  [[nodiscard]] std::string write_uri_ext(const std::string& uri) const;
  [[nodiscard]] std::string write_a16_creation_id() const;
  [[nodiscard]] std::string write_adec_decorative() const;
  [[nodiscard]] std::string write_a_hlink_click(uint32_t rel_index, const std::string& tip) const;
  [[nodiscard]] std::string write_a_pic_locks() const;
  [[nodiscard]] std::string write_c_nv_pic_pr() const;
  [[nodiscard]] std::string write_a_blip(uint32_t index) const;
  [[nodiscard]] std::string write_a_fill_rect() const;
  [[nodiscard]] std::string write_a_stretch() const;
  [[nodiscard]] std::string write_blip_fill(uint32_t index) const;
  [[nodiscard]] std::string write_a_ext(const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_a_off(const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_a_xfrm(const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_a_av_lst() const;
  [[nodiscard]] std::string write_a_prst_geom() const;
  [[nodiscard]] std::string write_sp_pr(const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_client_data() const;
  [[nodiscard]] std::string write_a_graphic_frame_locks() const;
  [[nodiscard]] std::string write_c_nv_graphic_frame_pr() const;
  [[nodiscard]] std::string write_nv_graphic_frame_pr(uint32_t index,
                                                      const std::optional<drawing_object_t>& drawing_object) const;
  [[nodiscard]] std::string write_xfrm_offset() const;
  [[nodiscard]] std::string write_xfrm_extension() const;
  [[nodiscard]] std::string write_xfrm() const;
  [[nodiscard]] std::string write_chart(uint32_t index) const;
  [[nodiscard]] std::string write_a_graphic_data(uint32_t index) const;
  [[nodiscard]] std::string write_a_graphic(uint32_t index) const;
  [[nodiscard]] std::string write_graphic_frame(uint32_t index, uint32_t rel_index,
                                                const std::optional<drawing_object_t>& drawing_object) const;
  [[nodiscard]] std::string write_ext(uint32_t cx, uint32_t cy) const;
  [[nodiscard]] std::string write_pos(int32_t x, int32_t y) const;

  bool embedded_                     = true;
  drawing_orientation_t orientation_ = drawing_orientation_t::LANDSCAPE;
  std::vector<drawing_object_t> drawing_objects_;
};

}

#endif
