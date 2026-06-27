/*
 * theme - A Xlsxwriter++ library for creating Excel XLSX theme files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
