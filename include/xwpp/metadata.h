/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * metadata - A Xlsxwriter++ library for creating Excel XLSX metadata files.
 *
 */

#ifndef XWPP_METADATA_H
#define XWPP_METADATA_H

#include "xwpp/common.h"

#include <cstdint>

namespace xwpp
{

/*
 * Struct to represent a metadata object.
 */
class metadata_t
{
public:
  metadata_t(bool has_dynamic_functions, bool has_embedded_images, uint32_t num_embedded_images);
  std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_metadata() const;
  [[nodiscard]] std::string write_metadata_types() const;
  [[nodiscard]] std::string write_cell_metadata_type() const;
  [[nodiscard]] std::string write_value_metadata_type() const;
  [[nodiscard]] std::string write_cell_future_metadata() const;
  [[nodiscard]] std::string write_xda_dynamic_array_properties() const;
  [[nodiscard]] std::string write_cell_ext() const;
  [[nodiscard]] std::string write_xlrd_rvb(uint32_t index) const;
  [[nodiscard]] std::string write_value_ext(uint32_t index) const;
  [[nodiscard]] std::string write_value_future_metadata() const;
  [[nodiscard]] std::string write_rc(uint8_t type, uint32_t index) const;
  [[nodiscard]] std::string write_cell_metadata() const;
  [[nodiscard]] std::string write_value_metadata() const;

  bool has_dynamic_functions_   = false;
  bool has_embedded_images_     = false;
  uint32_t num_embedded_images_ = 0;
};

}

#endif
