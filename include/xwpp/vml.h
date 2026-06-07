/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * vml - A Xlsxwriter++ library for creating Excel XLSX vml files.
 *
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
  vml_t(const std::string& vml_data_id_str, const std::vector<vml_obj_t>& comment_objs, uint32_t vml_shape_id,
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

  ///     uint8_t type;
  ///     struct lxw_comment_objs *button_objs;
  std::vector<vml_obj_t> comment_objs_;
  ///     struct lxw_comment_objs *image_objs;
  std::string vml_data_id_str_;
  uint32_t vml_shape_id_;
  comment_display_t comment_display_default_;
};

}

#endif
