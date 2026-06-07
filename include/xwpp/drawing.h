/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * drawing - A Xlsxwriter++ library for creating Excel XLSX drawing files.
 *
 */
#ifndef XWPP_DRAWING_H
#define XWPP_DRAWING_H

#include <cstdint>

namespace xwpp
{

/// STAILQ_HEAD(lxw_drawing_objects, lxw_drawing_object);

/// enum lxw_drawing_types {
///     LXW_DRAWING_NONE = 0,
///     LXW_DRAWING_IMAGE,
///     LXW_DRAWING_CHART,
///     LXW_DRAWING_SHAPE
/// };

/// enum image_types {
///     LXW_IMAGE_UNKNOWN = 0,
///     LXW_IMAGE_PNG,
///     LXW_IMAGE_JPEG,
///     LXW_IMAGE_BMP,
///     LXW_IMAGE_GIF
/// };

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
  ///     uint8_t type;
  uint8_t anchor_ = 0; // TODO Use object_position_t (but crossed inclusion for the moment)
  drawing_coords_t from_;
  drawing_coords_t to_;
  uint64_t col_absolute_;
  uint64_t row_absolute_;
  ///     uint32_t width;
  ///     uint32_t height;
  ///     uint8_t shape;
  ///     uint32_t rel_index;
  ///     uint32_t url_rel_index;
  ///     char *description;
  ///     char *tip;
  ///     uint8_t decorative;

  ///     STAILQ_ENTRY (lxw_drawing_object) list_pointers;
};

/*
 * Struct to represent a collection of drawings.
 */
/// typedef struct lxw_drawing {
///     FILE *file;
///     uint8_t embedded;
///     uint8_t orientation;
///     struct lxw_drawing_objects *drawing_objects;
/// } lxw_drawing;

/// lxw_drawing *lxw_drawing_new(void);
/// void lxw_drawing_free(lxw_drawing *drawing);
/// void lxw_drawing_assemble_xml_file(lxw_drawing *self);
/// void lxw_free_drawing_object(struct lxw_drawing_object *drawing_object);
/// void lxw_add_drawing_object(lxw_drawing *drawing,
///                             lxw_drawing_object *drawing_object);

/// STATIC void _drawing_xml_declaration(lxw_drawing *self);

}

#endif
