/*
 * rich_value_rel - A Xlsxwriter++ library for creating Excel XLSX rich_value_rel files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_RICH_VALUE_REL_H
#define XWPP_RICH_VALUE_REL_H

#include "xwpp/common.h"

#include <cstdint>
#include <string>

namespace xwpp
{

class rich_value_rel_t
{
public:
  rich_value_rel_t(uint32_t num_embedded_images);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_rel(uint32_t rel_index) const;
  [[nodiscard]] std::string write_rich_value_rels() const;

  uint32_t num_embedded_images_;
};

}

#endif
