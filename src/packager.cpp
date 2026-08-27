/*
 * packager - A library for assembling xml files into an Excel XLSX file.
 *
 * A class for writing the Excel XLSX Packager file.
 *
 * From Wikipedia: The Open Packaging Conventions (OPC) is a
 * container-file technology initially created by Microsoft to store
 * a combination of XML and non-XML files that together form a single
 * entity such as an Open XML Paper Specification (OpenXPS)
 * document. http://en.wikipedia.org/wiki/Open_Packaging_Conventions.
 *
 * At its simplest an Excel XLSX file contains the following elements::
 *
 *      ____ [Content_Types].xml
 *     |
 *     |____ docProps
 *     | |____ app.xml
 *     | |____ core.xml
 *     |
 *     |____ xl
 *     | |____ workbook.xml
 *     | |____ worksheets
 *     | | |____ sheet1.xml
 *     | |
 *     | |____ styles.xml
 *     | |
 *     | |____ theme
 *     | | |____ theme1.xml
 *     | |
 *     | |_____rels
 *     |   |____ workbook.xml.rels
 *     |
 *     |_____rels
 *       |____ .rels
 *
 * The packager class coordinates the classes that represent the
 * elements of the package and writes them into the XLSX file.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

// TODO XML are built in-memory. Should use temporary files to improve memory usage

#include "xwpp/packager.h"

#include "xwpp/app.h"
#include "xwpp/comment.h"
#include "xwpp/content_types.h"
#include "xwpp/core.h"
#include "xwpp/custom.h"
#include "xwpp/exception.h"
#include "xwpp/metadata.h"
#include "xwpp/relationships.h"
#include "xwpp/rich_value.h"
#include "xwpp/rich_value_rel.h"
#include "xwpp/rich_value_structure.h"
#include "xwpp/rich_value_types.h"
#include "xwpp/shared_strings.h"
#include "xwpp/styles.h"
#include "xwpp/table.h"
#include "xwpp/theme.h"
#include "xwpp/vml.h"
#include "xwpp/xmlwriter.h"

#include <algorithm>
#include <format>
#include <fstream>
#include <numeric>
#include <string>

namespace xwpp
{

packager_t::packager_t(std::string_view filename, bool use_zip64)
    // Initialize the zip_fileinfo struct to Jan 1 1980 like Excel.
    // TODO To set locally in create function.
    : zip_fileinfo_{
        .tmz_date =
            {
                .tm_sec = 0,
                .tm_min = 0,
                .tm_hour = 0,
                .tm_mday = 1,
                .tm_mon = 0,
                .tm_year = 1980,
            },
        .dosDate = 0,
        .internal_fa = 0,
        .external_fa = 0,

    }
    , filename_{filename}
    , use_zip64_{use_zip64}
{
}

// TODO const remove to allow prepare, to be refactored.
void packager_t::create_package(workbook_t& workbook)
{
  // TODO Use it as local variable (not data member) and use unique_ptr.
  zipfile_ = zipOpen(filename_.c_str(), 0);

  write_content_types_file(workbook);
  write_root_rels_file(workbook);
  write_workbook_rels_file(workbook);
  write_worksheet_files(workbook);
  write_chartsheet_files(workbook);
  write_workbook_file(workbook);
  write_chart_files(workbook);
  write_drawing_files(workbook);
  write_vml_files(workbook);
  write_comment_files(workbook);
  write_table_files(workbook);
  write_shared_strings_file(workbook);
  write_custom_file(workbook);
  write_theme_file();
  write_styles_file(workbook);
  write_worksheet_rels_file(workbook);
  write_chartsheet_rels_file(workbook);
  write_drawing_rels_file(workbook);
  write_image_files(workbook);
  add_vba_project(workbook);
  add_vba_project_signature(workbook);
  write_vba_project_rels_file(workbook);
  write_core_file(workbook);
  write_metadata_file(workbook);
  write_rich_value_file(workbook);
  write_rich_value_rel_file(workbook);
  write_rich_value_types_file(workbook);
  write_rich_value_structure_file(workbook);
  write_rich_value_rels_file(workbook);
  write_app_file(workbook);

  // TODO Manage errors (exception)
  /*zip_error =*/zipClose(zipfile_, nullptr);
  //     if (zip_error) {
  //         RETURN_ON_ZIP_ERROR(zip_error, LXW_ERROR_ZIP_CLOSE);
  //     }
}

void packager_t::write_content_types_file(const workbook_t& workbook)
{
  content_types_t content_types;
  uint32_t worksheet_index   = 1;
  uint32_t chartsheet_index  = 1;
  const size_t drawing_count = get_drawing_count(workbook);
  const size_t table_count   = get_table_count(workbook);

  if(workbook.has_png_)
  {
    content_types.add_default("png", "image/png");
  }

  if(workbook.has_jpeg_)
  {
    content_types.add_default("jpeg", "image/jpeg");
  }

  if(workbook.has_bmp_)
  {
    content_types.add_default("bmp", "image/bmp");
  }

  if(workbook.has_gif_)
  {
    content_types.add_default("gif", "image/gif");
  }

  if(!workbook.vba_project_.empty())
  {
    content_types.add_default("bin", "application/vnd.ms-office.vbaProject");
    content_types.add_override("/xl/workbook.xml", content_types_t::APP_MSEXCEL + "sheet.macroEnabled.main+xml");
  }
  else
  {
    content_types.add_override("/xl/workbook.xml", content_types_t::APP_DOCUMENT + "spreadsheetml.sheet.main+xml");
  }

  if(!workbook.vba_project_signature_.empty())
  {
    content_types.add_override("/xl/vbaProjectSignature.bin", "application/vnd.ms-office.vbaProjectSignature");
  }

  for(const auto* sheet: workbook.sheets_)
  {
    if(sheet->is_chartsheet())
    {
      content_types.add_chartsheet_name(std::format("/xl/chartsheets/sheet{}.xml", chartsheet_index));
      chartsheet_index++;
    }
    else
    {
      content_types.add_worksheet_name(std::format("/xl/worksheets/sheet{}.xml", worksheet_index));
      worksheet_index++;
    }
  }

  for(size_t index = 1; index <= workbook.ordered_charts_.size(); index++)
  {
    content_types.add_chart_name(std::format("/xl/charts/chart{}.xml", index));
  }

  for(size_t index = 1; index <= drawing_count; index++)
  {
    content_types.add_drawing_name(std::format("/xl/drawings/drawing{}.xml", index));
  }

  for(size_t index = 1; index <= table_count; index++)
  {
    content_types.add_table_name(std::format("/xl/tables/table{}.xml", index));
  }

  if(workbook.has_vml_)
  {
    content_types.add_vml_name();
  }

  for(size_t index = 1; index <= workbook.comment_count_; index++)
  {
    content_types.add_comment_name(std::format("/xl/comments{}.xml", index));
  }

  if(workbook.sst_.has_string())
  {
    content_types.add_shared_strings();
  }

  if(!workbook.custom_properties_.empty())
  {
    content_types.add_custom_properties();
  }

  if(workbook.has_metadata_)
  {
    content_types.add_metadata();
  }

  if(workbook.has_embedded_images_)
  {
    content_types.add_rich_value();
  }

  const std::string xml_data = content_types.assemble_xml_file();
  add_buffer_to_zip(xml_data, "[Content_Types].xml");
}

void packager_t::write_root_rels_file(const workbook_t& workbook)
{
  relationships_t relationships;
  relationships.add_document("/officeDocument", "xl/workbook.xml");
  relationships.add_package("/metadata/core-properties", "docProps/core.xml");
  relationships.add_document("/extended-properties", "docProps/app.xml");
  if(!workbook.custom_properties_.empty())
  {
    relationships.add_document("/custom-properties", "docProps/custom.xml");
  }
  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "_rels/.rels");
}

void packager_t::write_workbook_rels_file(const workbook_t& workbook)
{
  relationships_t relationships;
  uint32_t worksheet_index  = 1;
  uint32_t chartsheet_index = 1;

  for(const auto* sheet: workbook.sheets_)
  {
    if(sheet->is_chartsheet())
    {
      relationships.add_document("/chartsheet", std::format("chartsheets/sheet{}.xml", chartsheet_index));
      chartsheet_index++;
    }
    else
    {
      relationships.add_document("/worksheet", std::format("worksheets/sheet{}.xml", worksheet_index));
      worksheet_index++;
    }
  }

  relationships.add_document("/theme", "theme/theme1.xml");
  relationships.add_document("/styles", "styles.xml");

  if(workbook.sst_.has_string())
  {
    relationships.add_document("/sharedStrings", "sharedStrings.xml");
  }

  if(!workbook.vba_project_.empty())
  {
    relationships.add_ms_package("/vbaProject", "vbaProject.bin");
  }

  if(workbook.has_metadata_)
  {
    relationships.add_document("/sheetMetadata", "metadata.xml");
  }

  if(workbook.has_embedded_images_)
  {
    relationships.add_rich_value();
  }

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/_rels/workbook.xml.rels");
}

// cppcheck-suppress constParameterReference
void packager_t::write_worksheet_files(workbook_t& workbook)
{
  // Use ref to modify worksheet (add relations in external_hyperlinks_)
  for(size_t index = 1; auto& worksheet: workbook.worksheets_)
  {
    const std::string xml_data = worksheet.assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/worksheets/sheet{}.xml", index));
    index++;
  }
}

// cppcheck-suppress constParameterReference
void packager_t::write_chartsheet_files(workbook_t& workbook)
{
  for(size_t index = 1; auto& chartsheet: workbook.chartsheets_)
  {
    const std::string xml_data = chartsheet.assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/chartsheets/sheet{}.xml", index));
    index++;
  }
}

void packager_t::write_workbook_file(workbook_t& workbook)
{
  const std::string xml_data = workbook.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/workbook.xml");
}

void packager_t::write_styles_file(const workbook_t& workbook)
{
  // TODO Manage style in workbook
  style_t styles(workbook.font_count_, workbook.fill_count_, workbook.border_count_, workbook.num_format_count_,
                 workbook.has_comments_, workbook.used_xf_formats_, workbook.used_dxf_formats_);

  const std::string xml_data = styles.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/styles.xml");
}

void packager_t::write_shared_strings_file(const workbook_t& workbook)
{
  const shared_strings_t& shared_strings = workbook.sst_;

  // Skip the sharedStrings file if there are no shared strings.
  if(!shared_strings.has_string())
  {
    return;
  }

  const std::string xml_data = shared_strings.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/sharedStrings.xml");
}

void packager_t::write_theme_file()
{
  const std::string& xml_data = theme_t::assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/theme/theme1.xml");
}

void packager_t::write_app_file(const workbook_t& workbook)
{
  app_t app;

  if(!workbook.worksheets_.empty())
  {
    app.add_heading_pair("Worksheets", std::to_string(workbook.worksheets_.size()));
  }

  if(!workbook.chartsheets_.empty())
  {
    app.add_heading_pair("Charts", std::to_string(workbook.chartsheets_.size()));
  }

  // Two passes to have the same order of Excel: first worksheet and then chartsheet
  for(const auto& worksheet: workbook.worksheets_)
  {
    app.add_part_name(worksheet.get_sheet_name());
  }

  for(const auto& chartsheet: workbook.chartsheets_)
  {
    app.add_part_name(chartsheet.get_sheet_name());
  }

  // Add the Named Ranges parts.
  uint32_t named_range_count = 0;
  for(const auto& defined_name: workbook.defined_names_)
  {
    const bool has_range  = defined_name.formula_.find('!') != std::string::npos;
    const bool autofilter = defined_name.app_name_.find("_FilterDatabase") != std::string::npos;

    // Only store defined names with ranges (except for autofilters).
    if(has_range && !autofilter)
    {
      app.add_part_name(defined_name.app_name_);
      named_range_count++;
    }
  }

  // Add the Named Range heading pairs.
  if(named_range_count != 0)
  {
    app.add_heading_pair("Named Ranges", std::to_string(named_range_count));
  }

  // Set the app/doc properties.
  app.set_properties(workbook.properties_);

  app.set_doc_security(workbook.read_only_);

  const std::string xml_data = app.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/app.xml");
}

void packager_t::write_core_file(const workbook_t& workbook)
{
  const core_t core(workbook.properties_);

  const std::string xml_data = core.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/core.xml");
}

void packager_t::write_custom_file(const workbook_t& workbook)
{
  if(workbook.custom_properties_.empty())
  {
    return;
  }

  const custom_t custom(workbook.custom_properties_);

  const std::string xml_data = custom.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/custom.xml");
}

void packager_t::write_worksheet_rels_file(const workbook_t& workbook)
{
  uint32_t index = 0;

  for(const auto& worksheet: workbook.worksheets_)
  {
    relationships_t relationships;

    index++;

    if(worksheet.external_hyperlinks_.empty() && worksheet.external_drawing_links_.empty() &&
       worksheet.external_table_links_.empty() && !worksheet.external_vml_header_link_.has_value() &&
       !worksheet.external_vml_comment_link_.has_value() && !worksheet.external_background_link_.has_value() &&
       !worksheet.external_comment_link_.has_value())
    {
      continue;
    }

    for(const auto& [type, target, target_mode]: worksheet.external_hyperlinks_)
    {
      relationships.add_worksheet(type, target, target_mode);
    }

    for(const auto& [type, target, target_mode]: worksheet.external_drawing_links_)
    {
      relationships.add_worksheet(type, target, target_mode);
    }

    if(worksheet.external_vml_comment_link_.has_value())
    {
      const auto [type, target, target_mode] = worksheet.external_vml_comment_link_.value();
      relationships.add_worksheet(type, target, target_mode);
    }

    if(worksheet.external_vml_header_link_.has_value())
    {
      const auto [type, target, target_mode] = worksheet.external_vml_header_link_.value();
      relationships.add_worksheet(type, target, target_mode);
    }

    if(worksheet.external_background_link_.has_value())
    {
      const auto [type, target, target_mode] = worksheet.external_background_link_.value();
      relationships.add_worksheet(type, target, target_mode);
    }

    for(const auto& [type, target, target_mode]: worksheet.external_table_links_)
    {
      relationships.add_worksheet(type, target, target_mode);
    }

    if(worksheet.external_comment_link_.has_value())
    {
      const auto [type, target, target_mode] = worksheet.external_comment_link_.value();
      relationships.add_worksheet(type, target, target_mode);
    }

    const std::string xml_data = relationships.assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/worksheets/_rels/sheet{}.xml.rels", index));
  }
}

void packager_t::write_chartsheet_rels_file(const workbook_t& workbook)
{
  for(size_t index = 0; const auto& chartsheet: workbook.chartsheets_)
  {
    relationships_t relationships;

    index++;

    if(chartsheet.external_drawing_links_.empty())
    {
      continue;
    }

    for(const auto& [type, target, target_mode]: chartsheet.external_drawing_links_)
    {
      relationships.add_worksheet(type, target, target_mode);
    }

    const std::string xml_data = relationships.assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/chartsheets/_rels/sheet{}.xml.rels", index));
  }
}

void packager_t::write_vml_files(const workbook_t& workbook)
{
  uint32_t index = 1;

  for(const auto& worksheet: workbook.worksheets_)
  {
    if(!worksheet.has_vml_ && !worksheet.has_header_vml())
    {
      continue;
    }

    if(worksheet.has_vml_)
    {
      vml_t vml(worksheet.vml_data_id_str_, worksheet.comment_objs_, worksheet.button_objs_, worksheet.vml_shape_id_,
                worksheet.comment_display_default_);
      const std::string xml_data = vml.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/vmlDrawing{}.vml", index));
      index++;
    }

    if(worksheet.has_header_vml())
    {
      write_vml_drawing_rels_file(worksheet, index);
      vml_t vml(worksheet.vml_header_id_str_, worksheet.header_image_objs_, worksheet.vml_header_id_ * 1024);
      const std::string xml_data = vml.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/vmlDrawing{}.vml", index));
      index++;
    }
  }
}

void packager_t::write_comment_files(const workbook_t& workbook)
{
  uint32_t index = 1;

  for(const auto& worksheet: workbook.worksheets_)
  {
    if(!worksheet.has_comments_)
    {
      continue;
    }

    comment_t comment(worksheet.comment_objs_, worksheet.comment_author_);
    const std::string xml_data = comment.assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/comments{}.xml", index));
    index++;
  }
}

void packager_t::write_drawing_files(const workbook_t& workbook)
{
  for(size_t index = 1; const auto* sheet: workbook.sheets_)
  {
    if(sheet->drawing_.has_value())
    {
      const std::string xml_data = sheet->drawing_->assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/drawing{}.xml", index));
      index++;
    }
  }
}

void packager_t::write_drawing_rels_file(const workbook_t& workbook)
{
  for(size_t index = 1; const auto& sheet: workbook.sheets_)
  {
    if(!sheet->drawing_links_.empty())
    {
      relationships_t relationships;
      for(const auto& [type, target, target_mode]: sheet->drawing_links_)
      {
        relationships.add_worksheet(type, target, target_mode);
      }

      const std::string xml_data = relationships.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/_rels/drawing{}.xml.rels", index));
      index++;
    }
  }
}

void packager_t::write_image_files(const workbook_t& workbook)
{
  for(size_t index = 1; const auto& worksheet: workbook.worksheets_)
  {
    if(!worksheet.image_props_.empty() || worksheet.has_embedded_image())
    {
      for(const auto& object_props: worksheet.get_embedded_image_properties())
      {
        if(!object_props.is_duplicate_)
        {
          if(object_props.image_buffer_.empty())
          {
            // Read image.
            std::ifstream image_stream(object_props.filename_, std::ios::binary);
            const std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
            add_buffer_to_zip(buffer, std::format("xl/media/image{}.{}", index, object_props.extension_));
          }
          else
          {
            add_buffer_to_zip(object_props.image_buffer_,
                              std::format("xl/media/image{}.{}", index, object_props.extension_));
          }
          index++;
        }
      }

      for(const auto& object_props: worksheet.image_props_)
      {
        if(!object_props.is_duplicate_)
        {
          if(object_props.image_buffer_.empty())
          {
            // Read image.
            std::ifstream image_stream(object_props.filename_, std::ios::binary);
            const std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
            add_buffer_to_zip(buffer, std::format("xl/media/image{}.{}", index, object_props.extension_));
          }
          else
          {
            add_buffer_to_zip(object_props.image_buffer_,
                              std::format("xl/media/image{}.{}", index, object_props.extension_));
          }
          index++;
        }
      }
    }
  }
}

void packager_t::write_vml_drawing_rels_file(const worksheet_t& worksheet, uint32_t index)
{
  relationships_t relationships;

  for(const auto& [type, target, target_mode]: worksheet.vml_drawing_links_)
  {
    relationships.add_worksheet(type, target, target_mode);
  }

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, std::format("xl/drawings/_rels/vmlDrawing{}.vml.rels", index));
}

void packager_t::write_rich_value_rels_file(const workbook_t& workbook)
{
  if(!workbook.has_embedded_images_)
  {
    return;
  }

  relationships_t relationships;

  for(size_t index = 1; const auto& worksheet: workbook.worksheets_)
  {
    if(worksheet.has_embedded_image())
    {
      for(const auto& object_props: worksheet.get_embedded_image_properties())
      {
        if(!object_props.is_duplicate_)
        {
          relationships.add_document("/image", std::format("../media/image{}.{}", index, object_props.extension_));
          index++;
        }
      }
    }
  }

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/richData/_rels/richValueRel.xml.rels");
}

void packager_t::write_metadata_file(const workbook_t& workbook)
{
  if(workbook.has_metadata_)
  {
    const metadata_t metadata{workbook.has_dynamic_functions_, workbook.has_embedded_images_,
                              workbook.num_embedded_images_};

    const std::string xml_data = metadata.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/metadata.xml");
  }
}

void packager_t::write_rich_value_file(const workbook_t& workbook)
{
  if(!workbook.has_embedded_images_)
  {
    return;
  }

  const rich_value_t rich_value;

  const std::string xml_data = rich_value.assemble_xml_file(workbook);
  add_buffer_to_zip(xml_data, "xl/richData/rdrichvalue.xml");
}

void packager_t::write_rich_value_rel_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    const rich_value_rel_t rich_value_rel{workbook.num_embedded_images_};

    const std::string xml_data = rich_value_rel.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/richValueRel.xml");
  }
}

void packager_t::write_rich_value_structure_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    const rich_value_structure_t rich_value_structure{workbook.has_embedded_image_descriptions_};

    const std::string xml_data = rich_value_structure.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/rdrichvaluestructure.xml");
  }
}

void packager_t::write_rich_value_types_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    const std::string xml_data = rich_value_types_t::assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/rdRichValueTypes.xml");
  }
}

void packager_t::write_chart_files(const workbook_t& workbook)
{
  for(uint32_t index = 1; const auto& chart: workbook.ordered_charts_)
  {
    const std::string xml_data = chart->assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/charts/chart{}.xml", index));
    index++;
  }
}

void packager_t::write_table_files(const workbook_t& workbook)
{
  for(uint32_t index = 1; const auto& worksheet: workbook.worksheets_)
  {
    if(!worksheet.table_objs_.empty())
    {
      for(const auto& table_obj: worksheet.table_objs_)
      {
        const table_t table{table_obj};

        const std::string xml_data = table.assemble_xml_file();
        add_buffer_to_zip(xml_data, std::format("xl/tables/table{}.xml", index));
        index++;
      }
    }
  }
}

void packager_t::add_vba_project(const workbook_t& workbook)
{
  if(workbook.vba_project_.empty())
  {
    return;
  }

  std::ifstream vba_stream(workbook.vba_project_, std::ios::binary);
  const std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(vba_stream), {});
  add_buffer_to_zip(buffer, "xl/vbaProject.bin");
}

void packager_t::add_vba_project_signature(const workbook_t& workbook)
{
  if(workbook.vba_project_signature_.empty())
  {
    return;
  }

  // Check that the image file exists and can be opened.
  std::ifstream vba_signature_stream(workbook.vba_project_signature_, std::ios::binary);
  const std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(vba_signature_stream), {});
  add_buffer_to_zip(buffer, "xl/vbaProjectSignature.bin");
}

void packager_t::write_vba_project_rels_file(const workbook_t& workbook)
{
  if(workbook.vba_project_signature_.empty())
  {
    return;
  }

  relationships_t relationships;

  relationships.add_ms_package("/vbaProjectSignature", "vbaProjectSignature.bin");

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/_rels/vbaProject.bin.rels");
}

void packager_t::add_buffer_to_zip(std::string_view buffer, const std::string& filename)
{
  if(zipOpenNewFileInZip4_64(zipfile_, filename.c_str(), &zip_fileinfo_, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED,
                             Z_DEFAULT_COMPRESSION, 0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, nullptr, 0, 0, 0,
                             static_cast<int>(use_zip64_)) != ZIP_OK)
  {
    throw xwpp_exception_t(std::format("packager_t::add_buffer_to_zip(): error adding '{}' to zipfile", filename));
  }

  if(zipWriteInFileInZip(zipfile_, buffer.data(), static_cast<unsigned int>(buffer.size())) < 0)
  {
    throw xwpp_exception_t(
      std::format("packager_t::add_buffer_to_zip(): error in writing member '{}' to zipfile", filename));
  }

  if(zipCloseFileInZip(zipfile_) != ZIP_OK)
  {
    throw xwpp_exception_t(
      std::format("packager_t::add_buffer_to_zip(): error in closing member '{}' to zipfile", filename));
  }
}

void packager_t::add_buffer_to_zip(const std::vector<unsigned char>& buffer, const std::string& filename)
{
  if(zipOpenNewFileInZip4_64(zipfile_, filename.c_str(), &zip_fileinfo_, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED,
                             Z_DEFAULT_COMPRESSION, 0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, nullptr, 0, 0, 0,
                             static_cast<int>(use_zip64_)) != ZIP_OK)
  {
    throw xwpp_exception_t(
      std::format("packager_t::add_buffer_to_zip(): error adding member '{}' to zipfile", filename));
  }

  if(zipWriteInFileInZip(zipfile_, buffer.data(), static_cast<unsigned int>(buffer.size())) < 0)
  {
    throw xwpp_exception_t(
      std::format("packager_t::add_buffer_to_zip(): error in writing member '{}' to zipfile", filename));
  }

  if(zipCloseFileInZip(zipfile_) != ZIP_OK)
  {
    throw xwpp_exception_t(
      std::format("packager_t::add_buffer_to_zip(): error in closing member '{}' to zipfile", filename));
  }
}

size_t packager_t::get_drawing_count(const workbook_t& workbook) const
{
  return static_cast<size_t>(
    std::ranges::count_if(workbook.sheets_, [](const sheet_t* sheet) { return sheet->drawing_.has_value(); }));
}

size_t packager_t::get_table_count(const workbook_t& workbook) const
{
  return std::accumulate(std::begin(workbook.sheets_), std::end(workbook.sheets_), size_t{0},
                         [](size_t count, const sheet_t* sheet) { return count + sheet->get_table_count(); });
}

}
