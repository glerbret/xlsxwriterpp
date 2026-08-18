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
  vml_t(std::string vml_data_id_str, const std::vector<vml_obj_t>& image_objs, uint32_t vml_shape_id);
  vml_t(std::string vml_data_id_str, const std::vector<vml_obj_t>& comment_objs,
        const std::vector<vml_obj_t>& button_objs, uint32_t vml_shape_id, comment_display_t comment_display_default);

  [[nodiscard]] std::string assemble_xml_file();

private:
  [[nodiscard]] std::string write_shapelayout() const;
  [[nodiscard]] std::string write_idmap() const;
  [[nodiscard]] std::string write_comment_shape(uint32_t vml_shape_id, uint32_t z_index, vml_obj_t& vml_obj) const;
  [[nodiscard]] static std::string write_comment_client_data(const vml_obj_t& vml_obj);
  [[nodiscard]] static std::string write_image_shape(uint32_t vml_shape_id, uint32_t z_index,
                                                     const vml_obj_t& image_obj);
  [[nodiscard]] static std::string write_button_shape(uint32_t vml_shape_id, uint32_t z_index,
                                                      const vml_obj_t& vml_obj);

  std::vector<vml_obj_t> button_objs_;
  std::vector<vml_obj_t> comment_objs_;
  std::vector<vml_obj_t> image_objs_;
  std::string vml_data_id_str_;
  uint32_t vml_shape_id_                     = 0;
  comment_display_t comment_display_default_ = comment_display_t::DEFAULT;
};

}

#endif
