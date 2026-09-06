/*
 * comment - A library for creating Excel XLSX comment files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/comment.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>
#include <tuple>
#include <vector>

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_comments()
{
  return xml_start_tag("comments", {
                                     {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
  });
}

[[nodiscard]] std::string write_author(const std::string& author)
{
  return xml_data_element("author", author);
}

[[nodiscard]] std::string write_text_t(const vml_obj_t& comment)
{
  return xml_data_element("t", comment.text_);
}

[[nodiscard]] std::string write_family(const vml_obj_t& comment)
{
  return xml_empty_tag("family", {
                                   {"val", std::to_string(comment.font_family_)}
  });
}

[[nodiscard]] std::string write_r_font(const vml_obj_t& comment)
{
  attributes_t attributes;

  if(!comment.font_name_.empty())
  {
    attributes.add_attribute("val", comment.font_name_);
  }
  else
  {
    attributes.add_attribute("val", "Tahoma");
  }

  return xml_empty_tag("rFont", attributes);
}

[[nodiscard]] std::string write_color()
{
  return xml_empty_tag("color", {
                                  {"indexed", "81"}
  });
}

[[nodiscard]] std::string write_sz(const vml_obj_t& comment)
{
  return xml_empty_tag("sz", {
                               {"val", std::format("{}", comment.font_size_)}
  });
}

[[nodiscard]] std::string write_r_pr(const vml_obj_t& comment)
{
  std::string xml_data = xml_start_tag("rPr");
  xml_data += write_sz(comment);
  xml_data += write_color();
  xml_data += write_r_font(comment);
  xml_data += write_family(comment);
  xml_data += xml_end_tag("rPr");

  return xml_data;
}

[[nodiscard]] std::string write_r(const vml_obj_t& comment)
{
  std::string xml_data = xml_start_tag("r");
  xml_data += write_r_pr(comment);
  xml_data += write_text_t(comment);
  xml_data += xml_end_tag("r");

  return xml_data;
}

[[nodiscard]] std::string write_text(const vml_obj_t& comment)
{
  std::string xml_data = xml_start_tag("text");
  xml_data += write_r(comment);
  xml_data += xml_end_tag("text");

  return xml_data;
}

[[nodiscard]] std::string write_comment(const vml_obj_t& comment)
{
  const std::string ref = rowcol_to_cell(comment.row_num_, comment.col_num_);
  std::string xml_data  = xml_start_tag("comment", {
                                                    {"ref",      ref                               },
                                                    {"authorId", std::to_string(comment.author_id_)},
  });
  xml_data += write_text(comment);
  xml_data += xml_end_tag("comment");

  return xml_data;
}

}

comment_t::comment_t(const std::vector<vml_obj_t>& comment_objs, std::string comment_author)
  : comment_objs_{comment_objs}
  , comment_author_{std::move(comment_author)}
{
}

std::string comment_t::assemble_xml_file()
{
  std::string xml_data = xml_declaration();
  xml_data += write_comments();
  xml_data += write_authors();
  xml_data += write_comment_list();
  xml_data += xml_end_tag("comments");

  return xml_data;
}

bool comment_t::check_author(const std::string& author) const
{
  if(author.empty())
  {
    return true;
  }

  return author_ids_.find(author) != std::end(author_ids_);
}

uint32_t comment_t::get_author_index(const std::string& author)
{
  if(const auto it = author_ids_.find(author); it != std::end(author_ids_))
  {
    return it->second;
  }

  author_ids_[author] = author_id_;
  author_id_++;

  return author_id_ - 1;
}

std::string comment_t::write_authors()
{
  std::string xml_data = xml_start_tag("authors");

  // Set the default author (from worksheet_set_comments_author()).
  if(!comment_author_.empty())
  {
    (void)get_author_index(comment_author_);
    xml_data += write_author(comment_author_);
  }
  else
  {
    (void)get_author_index("");
    xml_data += write_author("");
  }

  for(auto& comment: comment_objs_)
  {
    if(!comment.author_.empty())
    {
      if(!check_author(comment.author_))
      {
        xml_data += write_author(comment.author_);
      }

      comment.author_id_ = get_author_index(comment.author_);
    }
  }

  xml_data += xml_end_tag("authors");

  return xml_data;
}

std::string comment_t::write_comment_list() const
{
  std::string xml_data = xml_start_tag("commentList");
  for(const auto& comment: comment_objs_)
  {
    xml_data += write_comment(comment);
  }
  xml_data += xml_end_tag("commentList");

  return xml_data;
}

}
