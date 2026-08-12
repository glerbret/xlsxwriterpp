/*
 * rich_value_structure - A Xlsxwriter++ library for creating Excel XLSX rich_value_structure files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_RICH_VALUE_STRUCTURE_H
#define XWPP_RICH_VALUE_STRUCTURE_H

#include "xwpp/common.h"

#include <cstdint>

namespace xwpp
{

class rich_value_structure_t
{
public:
  explicit rich_value_structure_t(bool has_embedded_image_descriptions);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_rv_structures() const;
  [[nodiscard]] std::string write_s() const;
  [[nodiscard]] std::string write_k(const std::string& name, const std::string& type) const;

  bool has_embedded_image_descriptions_;
};

}

#endif
