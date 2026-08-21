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
  uint32_t col_num_  = 0;
  uint32_t row_num_  = 0;
  double col_offset_ = 0.;
  double row_offset_ = 0.;
};

// Object to represent the properties of a drawing.
struct drawing_object_t
{
  drawing_types_t type_ = drawing_types_t::NONE;
  uint8_t anchor_       = 0; // TODO Use object_position_t (but crossed inclusion for the moment)
  drawing_coords_t from_;
  drawing_coords_t to_;
  uint64_t col_absolute_  = 0;
  uint64_t row_absolute_  = 0;
  uint32_t width_         = 0;
  uint32_t height_        = 0;
  // TODO ?     uint8_t shape;
  uint32_t rel_index_     = 0;
  uint32_t url_rel_index_ = 0;
  std::string description_;
  std::string tip_;
  bool decorative_ = false;
};

class drawing_t
{
public:
  void add_drawing_object(const drawing_object_t& drawing_object);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  friend class worksheet_t;

  [[nodiscard]] std::string write_absolute_anchor(uint32_t frame_index) const;
  [[nodiscard]] std::string write_two_cell_anchor(uint32_t index, const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_c_nv_graphic_frame_pr() const;
  [[nodiscard]] std::string write_nv_graphic_frame_pr(uint32_t index,
                                                      const std::optional<drawing_object_t>& drawing_object) const;
  [[nodiscard]] std::string write_graphic_frame(uint32_t index, uint32_t rel_index,
                                                const std::optional<drawing_object_t>& drawing_object) const;

  bool embedded_                     = true;
  drawing_orientation_t orientation_ = drawing_orientation_t::LANDSCAPE;
  std::vector<drawing_object_t> drawing_objects_;
};

}

#endif
