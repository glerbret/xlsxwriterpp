/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/core.h"

#include <boost/test/unit_test.hpp>

#include <chrono>

using namespace std::literals::chrono_literals;

#ifdef _WIN32
#define timegm _mkgmtime
#define strdup _strdup
#endif

BOOST_AUTO_TEST_SUITE(core)

BOOST_AUTO_TEST_CASE(assemble_xml_file_1)
{
  const std::string expected =
      // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        "<dc:creator>A User</dc:creator>"
        "<cp:lastModifiedBy/>"
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2010-01-01T00:00:00Z</dcterms:created>"
      "</cp:coreProperties>";
  // clang-format on

  const xwpp::doc_properties_t properties{
      .author_  = "A User",
      .created_ = std::chrono::sys_days{2010y / std::chrono::January / 01d} + 0h + 0min + 0s,
  };
  const xwpp::core_t core(properties);

  BOOST_CHECK_EQUAL(expected, core.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(assemble_xml_file_2)
{
  const std::string expected =
      // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
        "<dc:title>This is an example spreadsheet</dc:title>"
        "<dc:subject>With document properties</dc:subject>"
        "<dc:creator>A Person</dc:creator>"
        "<cp:keywords>Sample, Example, Properties</cp:keywords>"
        "<dc:description>Created with Xlsxwriter++</dc:description>"
        "<cp:lastModifiedBy/>"
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2011-04-06T19:45:15Z</dcterms:created>"
        "<cp:category>Example spreadsheets</cp:category>"
        "<cp:contentStatus>Quo</cp:contentStatus>"
      "</cp:coreProperties>";
  // clang-format on

  const xwpp::doc_properties_t properties{
      .title_    = "This is an example spreadsheet",
      .subject_  = "With document properties",
      .author_   = "A Person",
      .category_ = "Example spreadsheets",
      .keywords_ = "Sample, Example, Properties",
      .comments_ = "Created with Xlsxwriter++",
      .status_   = "Quo",
      .created_  = std::chrono::sys_days{2011y / std::chrono::April / 06d} + 19h + 45min + 15s,
  };
  const xwpp::core_t core(properties);

  BOOST_CHECK_EQUAL(expected, core.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
