/*
 * rich_value - A Xlsxwriter++ library for creating Excel XLSX rich_value files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_RICH_VALUE_H
#define XWPP_RICH_VALUE_H

#include "xwpp/common.h"
#include "xwpp/workbook.h"

namespace xwpp
{

class rich_value_t
{
public:
  [[nodiscard]] std::string assemble_xml_file(const workbook_t& workbook) const;

private:
  [[nodiscard]] static std::string write_rv_data(const workbook_t& workbook);
  [[nodiscard]] std::string write_images(const workbook_t& workbook) const;
};

}

#endif
