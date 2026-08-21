/*
 * metadata - A library for creating Excel XLSX metadata files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/metadata.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_cell_metadata_type()
{
  return xml_empty_tag("metadataType", {
                                         {"name",                "XLDAPR"},
                                         {"minSupportedVersion", "120000"},
                                         {"copy",                "1"     },
                                         {"pasteAll",            "1"     },
                                         {"pasteValues",         "1"     },
                                         {"merge",               "1"     },
                                         {"splitFirst",          "1"     },
                                         {"rowColShift",         "1"     },
                                         {"clearFormats",        "1"     },
                                         {"clearComments",       "1"     },
                                         {"assign",              "1"     },
                                         {"coerce",              "1"     },
                                         {"cellMeta",            "1"     },
  });
}

[[nodiscard]] std::string write_value_metadata_type()
{
  return xml_empty_tag("metadataType", {
                                         {"name",                "XLRICHVALUE"},
                                         {"minSupportedVersion", "120000"     },
                                         {"copy",                "1"          },
                                         {"pasteAll",            "1"          },
                                         {"pasteValues",         "1"          },
                                         {"merge",               "1"          },
                                         {"splitFirst",          "1"          },
                                         {"rowColShift",         "1"          },
                                         {"clearFormats",        "1"          },
                                         {"clearComments",       "1"          },
                                         {"assign",              "1"          },
                                         {"coerce",              "1"          },
  });
}

[[nodiscard]] std::string write_xda_dynamic_array_properties()
{
  return xml_empty_tag("xda:dynamicArrayProperties", {
                                                       {"fDynamic",   "1"},
                                                       {"fCollapsed", "0"},
  });
}

[[nodiscard]] std::string write_xlrd_rvb(uint32_t index)
{
  return xml_empty_tag("xlrd:rvb", {
                                     {"i", std::to_string(index)}
  });
}

[[nodiscard]] std::string write_rc(uint8_t type, uint32_t index)
{
  return xml_empty_tag("rc", {
                               {"t", std::to_string(type) },
                               {"v", std::to_string(index)},
  });
}

[[nodiscard]] std::string write_cell_ext()
{
  std::string xml_data = xml_start_tag("ext", {
                                                {"uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"}
  });

  xml_data += write_xda_dynamic_array_properties();
  xml_data += xml_end_tag("ext");

  return xml_data;
}

[[nodiscard]] std::string write_value_ext(uint32_t index)
{
  std::string xml_data = xml_start_tag("ext", {
                                                {"uri", "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"}
  });

  xml_data += write_xlrd_rvb(index);
  xml_data += xml_end_tag("ext");

  return xml_data;
}

[[nodiscard]] std::string write_cell_metadata()
{
  std::string xml_data = xml_start_tag("cellMetadata", {
                                                         {"count", "1"}
  });
  xml_data += xml_start_tag("bk");
  xml_data += write_rc(1, 0);
  xml_data += xml_end_tag("bk");
  xml_data += xml_end_tag("cellMetadata");

  return xml_data;
}

[[nodiscard]] std::string write_cell_future_metadata()
{
  std::string xml_data = xml_start_tag("futureMetadata", {
                                                           {"name",  "XLDAPR"},
                                                           {"count", "1"     },
  });

  xml_data += xml_start_tag("bk");
  xml_data += xml_start_tag("extLst");
  xml_data += write_cell_ext();
  xml_data += xml_end_tag("extLst");
  xml_data += xml_end_tag("bk");
  xml_data += xml_end_tag("futureMetadata");

  return xml_data;
}

}

metadata_t::metadata_t(bool has_dynamic_functions, bool has_embedded_images, uint32_t num_embedded_images)
  : has_dynamic_functions_{has_dynamic_functions}
  , has_embedded_images_{has_embedded_images}
  , num_embedded_images_{num_embedded_images}
{
}

std::string metadata_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_metadata();
  xml_data += write_metadata_types();
  if(has_dynamic_functions_)
  {
    xml_data += write_cell_future_metadata();
  }
  if(has_embedded_images_)
  {
    xml_data += write_value_future_metadata();
  }
  if(has_dynamic_functions_)
  {
    xml_data += write_cell_metadata();
  }
  if(has_embedded_images_)
  {
    xml_data += write_value_metadata();
  }
  xml_data += xml_end_tag("metadata");

  return xml_data;
}

std::string metadata_t::write_metadata() const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
    {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
  };

  if(has_embedded_images_)
  {
    attributes.emplace_back("xmlns:xlrd", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata");
  }

  if(has_dynamic_functions_)
  {
    attributes.emplace_back("xmlns:xda", "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray");
  }

  return xml_start_tag("metadata", attributes);
}

std::string metadata_t::write_metadata_types() const
{
  uint8_t count = 0;

  if(has_dynamic_functions_)
  {
    count++;
  }

  if(has_embedded_images_)
  {
    count++;
  }

  std::string xml_data = xml_start_tag("metadataTypes", {
                                                          {"count", std::to_string(count)}
  });

  if(has_dynamic_functions_)
  {
    xml_data += write_cell_metadata_type();
  }
  if(has_embedded_images_)
  {
    xml_data += write_value_metadata_type();
  }
  xml_data += xml_end_tag("metadataTypes");

  return xml_data;
}

std::string metadata_t::write_value_future_metadata() const
{
  std::string xml_data = xml_start_tag("futureMetadata", {
                                                           {"name",  "XLRICHVALUE"                       },
                                                           {"count", std::to_string(num_embedded_images_)},
  });

  for(uint32_t i = 0; i < num_embedded_images_; i++)
  {
    xml_data += xml_start_tag("bk");
    xml_data += xml_start_tag("extLst");
    xml_data += write_value_ext(i);
    xml_data += xml_end_tag("extLst");
    xml_data += xml_end_tag("bk");
  }

  xml_data += xml_end_tag("futureMetadata");

  return xml_data;
}

std::string metadata_t::write_value_metadata() const
{
  uint8_t type = 1;

  if(has_dynamic_functions_)
  {
    type = 2;
  }

  std::string xml_data = xml_start_tag("valueMetadata", {
                                                          {"count", std::to_string(num_embedded_images_)}
  });
  for(uint32_t i = 0; i < num_embedded_images_; i++)
  {
    xml_data += xml_start_tag("bk");
    xml_data += write_rc(type, i);
    xml_data += xml_end_tag("bk");
  }
  xml_data += xml_end_tag("valueMetadata");

  return xml_data;
}

}
