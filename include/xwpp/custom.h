/*
 * custom - A Xlsxwriter++ library for creating Excel custom property files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_CUSTOM_H
#define XWPP_CUSTOM_H

#include "xwpp/common.h"

#include <string>
#include <vector>

namespace xwpp
{

class custom_t
{
public:
  explicit custom_t(const std::vector<custom_property_t>& custom_properties);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_custom_properties() const;

  std::vector<custom_property_t> custom_properties_;
};

}

#endif
