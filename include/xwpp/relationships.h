/*
 * relationships - A Xlsxwriter++ library for creating Excel XLSX
 *                 relationships files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_RELATIONSHIPS_H
#define XWPP_RELATIONSHIPS_H

#include <cstdint>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

namespace xwpp
{

class relationships_t
{
public:
  void add_document(std::string_view type, std::string_view target);
  void add_package(std::string_view type, std::string_view target);
  void add_worksheet(std::string_view type, std::string_view target, std::string_view target_mode);
  void add_rich_value();
  void add_ms_package(std::string_view type, std::string_view target);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  void add(std::string_view schema, std::string_view type, std::string_view target, std::string_view target_mode);
  void add(std::string_view schema, std::string_view type, std::string_view target);

  /// uint32_t rel_id = 0;
  std::vector<std::tuple<std::string, std::string, std::string>> relationships_;
};

}

#endif
