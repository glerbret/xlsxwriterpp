/*
 * vml - A Xlsxwriter++ library for creating Excel XLSX vml files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_VML_H
#define XWPP_VML_H

#include "xwpp/comment.h"

#include <cstdint>
#include <string>
#include <vector>

namespace xwpp
{

class vml_t
{
public:
  vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& comment_objs, uint32_t vml_shape_id);
  vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& image_objs, uint32_t vml_shape_id,
        comment_display_t comment_display_default);

  [[nodiscard]] std::string assemble_xml_file();

private:
  [[nodiscard]] std::string write_xml_namespace() const;
  [[nodiscard]] std::string write_shapelayout() const;
  [[nodiscard]] std::string write_idmap() const;
  [[nodiscard]] std::string write_comment_shapetype() const;
  [[nodiscard]] std::string write_stroke() const;
  [[nodiscard]] std::string write_comment_path(bool has_gradient, const std::string& type) const;
  [[nodiscard]] std::string write_comment_shape(uint32_t vml_shape_id, uint32_t z_index, vml_obj_t& vml_obj) const;
  [[nodiscard]] std::string write_comment_fill() const;
  [[nodiscard]] std::string write_shadow() const;
  [[nodiscard]] std::string write_comment_textbox() const;
  [[nodiscard]] std::string write_comment_div() const;
  [[nodiscard]] std::string write_comment_client_data(const vml_obj_t& vml_obj) const;
  [[nodiscard]] std::string write_move_with_cells() const;
  [[nodiscard]] std::string write_size_with_cells() const;
  [[nodiscard]] std::string write_anchor(const vml_obj_t& vml_obj) const;
  [[nodiscard]] std::string write_auto_fill() const;
  [[nodiscard]] std::string write_row(const vml_obj_t& vml_obj) const;
  [[nodiscard]] std::string write_column(const vml_obj_t& vml_obj) const;
  [[nodiscard]] std::string write_visible() const;
  [[nodiscard]] std::string write_image_shapetype() const;
  [[nodiscard]] std::string write_formulas() const;
  [[nodiscard]] std::string write_formula(const std::string& equation) const;
  [[nodiscard]] std::string write_image_path() const;
  [[nodiscard]] std::string write_aspect_ratio_lock() const;
  [[nodiscard]] std::string write_image_shape(uint32_t vml_shape_id, uint32_t z_index,
                                              const vml_obj_t& image_obj) const;
  [[nodiscard]] std::string write_imagedata(uint32_t rel_index, const std::string& name) const;
  [[nodiscard]] std::string write_rotation_lock() const;

  ///     uint8_t type;
  ///     struct lxw_comment_objs *button_objs;
  std::vector<vml_obj_t> comment_objs_;
  std::vector<vml_obj_t> image_objs_;
  std::string vml_data_id_str_;
  uint32_t vml_shape_id_;
  comment_display_t comment_display_default_;
};

}

#endif
