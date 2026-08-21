/*
 * rich_value_types - A Xlsxwriter++ library for creating Excel XLSX rich_value_types files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_RICH_VALUE_TYPES_H
#define XWPP_RICH_VALUE_TYPES_H

#include "xwpp/common.h"

#include <cstdint>

namespace xwpp
{

class rich_value_types_t
{
public:
  [[nodiscard]] static std::string assemble_xml_file();
};

}

#endif
