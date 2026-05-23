/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * theme - A Xlsxwriter++ library for creating Excel XLSX theme files.
 *
 */

#ifndef XWPP_THEME_H
#define XWPP_THEME_H

#include <string>

namespace xwpp
{

class theme_t
{
public:
  [[nodiscard]] std::string assemble_xml_file() const;

private:
  static const std::string THEME;
};

}

#endif
