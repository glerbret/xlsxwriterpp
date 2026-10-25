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
  uint32_t col_num{0};
  uint32_t row_num{0};
  double col_offset{0.};
  double row_offset{0.};
};

// Object to represent the properties of a drawing.
struct drawing_object_t
{
  drawing_types_t type{drawing_types_t::NONE};
  object_position_t anchor{object_position_t::DEFAULT};
  drawing_coords_t from;
  drawing_coords_t to;
  uint64_t col_absolute{0};
  uint64_t row_absolute{0};
  uint32_t width{0};
  uint32_t height{0};
  uint32_t rel_index{0};
  uint32_t url_rel_index{0};
  std::string description;
  std::string tip;
  bool decorative{false};
};

class drawing_t
{
public:
  void add_drawing_object(const drawing_object_t& drawing_object);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  friend class sheet_t;
  friend class chartsheet_t;
  friend class worksheet_t;

  [[nodiscard]] std::string write_absolute_anchor(uint32_t frame_index) const;
  [[nodiscard]] std::string write_two_cell_anchor(uint32_t index, const drawing_object_t& drawing_object) const;
  [[nodiscard]] std::string write_c_nv_graphic_frame_pr() const;
  [[nodiscard]] std::string write_nv_graphic_frame_pr(uint32_t index,
                                                      const std::optional<drawing_object_t>& drawing_object) const;
  [[nodiscard]] std::string write_graphic_frame(uint32_t index, uint32_t rel_index,
                                                const std::optional<drawing_object_t>& drawing_object) const;

  bool embedded_{true};
  drawing_orientation_t orientation_{drawing_orientation_t::LANDSCAPE};
  std::vector<drawing_object_t> drawing_objects_;
};

}

#endif
