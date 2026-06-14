/*
 * Copyright 2026, Grégory Lerbret
 *
 * rich_value - A Xlsxwriter++ library for creating Excel XLSX rich_value files.
 *
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
  std::string assemble_xml_file(const workbook_t& workbook) const;

private:
  [[nodiscard]] std::string write_rv_data(const workbook_t& workbook) const;
  [[nodiscard]] std::string write_images(const workbook_t& workbook) const;
  [[nodiscard]] std::string write_rv() const;
  [[nodiscard]] std::string write_v(const std::string& value) const;
};

}

#endif
