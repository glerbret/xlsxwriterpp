/*
 * app - A Xlsxwriter++ library for creating Excel XLSX app files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_APP_H
#define XWPP_APP_H

#include "common.h"
#include "workbook.h"

#include <cstdint>
#include <string>
#include <vector>

namespace xwpp
{

struct heading_pair_t
{
  std::string key_;
  std::string value_;
};

struct part_name_t
{
  std::string name_;
};

class app_t
{
public:
  void add_part_name(const std::string& name);
  void add_heading_pair(const std::string& key, const std::string& value);
  void set_properties(const doc_properties_t& properties);
  void set_doc_security(uint8_t doc_security);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_properties() const;
  [[nodiscard]] std::string write_application() const;
  [[nodiscard]] std::string write_doc_security() const;
  [[nodiscard]] std::string write_scale_crop() const;
  [[nodiscard]] std::string write_heading_pairs() const;
  [[nodiscard]] std::string write_titles_of_parts() const;
  [[nodiscard]] std::string write_manager() const;
  [[nodiscard]] std::string write_company() const;
  [[nodiscard]] std::string write_links_up_to_date() const;
  [[nodiscard]] std::string write_shared_doc() const;
  [[nodiscard]] std::string write_hyperlink_base() const;
  [[nodiscard]] std::string write_hyperlinks_changed() const;
  [[nodiscard]] std::string write_app_version() const;
  [[nodiscard]] std::string write_vt_vector_heading_pairs() const;
  [[nodiscard]] std::string write_vt_vector_lpstr_named_parts() const;
  [[nodiscard]] std::string write_vt_lpstr(const std::string& str) const;
  [[nodiscard]] std::string write_vt_variant(const std::string& key, const std::string& value) const;
  [[nodiscard]] std::string write_vt_i4(const std::string& value) const;

  std::vector<heading_pair_t> heading_pairs_;
  std::vector<part_name_t> part_names_;
  doc_properties_t properties_;
  uint8_t doc_security_ = 0;
};

}

#endif
