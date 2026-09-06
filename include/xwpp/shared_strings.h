/*
 * shared_strings - A Xlsxwriter++ library for creating Excel XLSX
 *                  sst files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
  uint32_t index_{0};
  std::string string_;
  bool is_rich_string_{false};
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

  uint32_t string_count_{0};
  uint32_t unique_count_{0};

  std::vector<shared_strings_element_t> order_list_;
  std::map<std::string, size_t, std::less<>> rb_tree_;
};

}

#endif
