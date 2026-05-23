/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * shared_strings - A Xlsxwriter++ library for creating Excel XLSX
 *                  sst files.
 *
 */

#ifndef XWPP_SHARED_STRINGS_H
#define XWPP_SHARED_STRINGS_H

#include <cstdint>
#include <map>
#include <string>
#include <vector>

namespace xwpp
{

struct shared_strings_element_t
{
  uint32_t index_ = 0;
  std::string string_;
  bool is_rich_string_;
};

class shared_strings_t
{
public:
  [[nodiscard]] bool has_string() const;
  [[nodiscard]] std::string assemble_xml_file() const;
  [[nodiscard]] shared_strings_element_t get_index(const std::string& str, bool is_rich_string);

private:
  [[nodiscard]] std::string write_sst() const;
  [[nodiscard]] std::string write_sst_strings() const;
  [[nodiscard]] std::string write_si(const std::string& str) const;
  [[nodiscard]] std::string write_rich_si(const std::string& str) const;
  [[nodiscard]] std::string write_t(const std::string& str) const;

  uint32_t string_count_ = 0;
  uint32_t unique_count_ = 0;

  std::vector<shared_strings_element_t> order_list_;
  std::map<std::string, size_t> rb_tree_;
};

}

#endif
