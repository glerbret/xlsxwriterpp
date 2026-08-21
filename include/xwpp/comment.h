/*
 * comment - A Xlsxwriter++ library for creating Excel XLSX comment files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_COMMENT_H
#define XWPP_COMMENT_H

#include <cstdint>
#include <map>
#include <string>

#include "xwpp/worksheet.h"

namespace xwpp
{

class comment_t
{
public:
  comment_t(const std::vector<vml_obj_t>& comment_objs, std::string comment_author);

  [[nodiscard]] std::string assemble_xml_file();

private:
  [[nodiscard]] bool check_author(const std::string& author) const;
  [[nodiscard]] uint32_t get_author_index(const std::string& author);
  [[nodiscard]] std::string write_authors();
  [[nodiscard]] std::string write_comment_list() const;

  std::vector<vml_obj_t> comment_objs_;
  std::map<std::string, uint32_t> author_ids_;
  std::string comment_author_;
  uint32_t author_id_ = 0;
};

}

#endif
