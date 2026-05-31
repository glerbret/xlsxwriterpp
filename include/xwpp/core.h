/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * core - A Xlsxwriter++ library for creating Excel XLSX core files.
 *
 */

#ifndef XWPP_CORE_H
#define XWPP_CORE_H

#include "xwpp/workbook.h"

#include <string>

namespace xwpp
{

class core_t
{
public:
  [[nodiscard]] std::string assemble_xml_file() const;
  void set_properties(const doc_properties_t& properties);

private:
  [[nodiscard]] std::string write_cp_core_properties() const;

  [[nodiscard]] std::string write_dc_title() const;
  [[nodiscard]] std::string write_dc_subject() const;
  [[nodiscard]] std::string write_dc_creator() const;
  [[nodiscard]] std::string write_cp_keywords() const;
  [[nodiscard]] std::string write_dc_description() const;
  [[nodiscard]] std::string write_cp_last_modified_by() const;
  [[nodiscard]] std::string write_dcterms_created() const;
  [[nodiscard]] std::string write_dcterms_modified() const;
  [[nodiscard]] std::string write_cp_category() const;
  [[nodiscard]] std::string write_cp_content_status() const;

  doc_properties_t properties_;
};

}

#endif
