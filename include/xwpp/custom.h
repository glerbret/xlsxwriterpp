/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * custom - A Xlsxwriter++ library for creating Excel custom property files.
 *
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

  [[nodiscard]] std::string assemble_xml_file();

private:
  [[nodiscard]] std::string write_custom_properties() const;
  [[nodiscard]] std::string write_custom_property(size_t pid, const custom_property_t& property) const;
  [[nodiscard]] std::string write(const std::string& value) const;
  [[nodiscard]] std::string write(double value) const;
  [[nodiscard]] std::string write(int32_t value) const;
  [[nodiscard]] std::string write(bool value) const;
  [[nodiscard]] std::string write(const std::chrono::system_clock::time_point& value) const;

  std::vector<custom_property_t> custom_properties_;
};

}

#endif
