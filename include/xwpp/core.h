/*
 * core - A Xlsxwriter++ library for creating Excel XLSX core files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
  explicit core_t(doc_properties_t properties);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
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
