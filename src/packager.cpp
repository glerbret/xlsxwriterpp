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
#include "xwpp/theme.h"
#include "xwpp/vml.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <fstream>
#include <string>

#include <iostream>

namespace xwpp
{

/// STATIC lxw_error _add_file_to_zip(lxw_packager *self, FILE *file,
///                                   const char *filename);

/// STATIC lxw_error _add_buffer_to_zip(lxw_packager *self, const char *buffer,
///                                     size_t buffer_size, const char *filename);

/// STATIC lxw_error _add_to_zip(lxw_packager *self, FILE *file,
///                              char **buffer, size_t *buffer_size,
///                              const char *filename);

packager_t::packager_t(std::string_view filename /*, const char *tmpdir, uint8_t use_zip64*/)
    // Initialize the zip_fileinfo struct to Jan 1 1980 like Excel.
    // TODO To set locally in create function
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
{
  /// packager->tmpdir = tmpdir;
  /// packager->buffer_size = LXW_ZIP_BUFFER_SIZE;
  /// packager->use_zip64 = use_zip64;
}

void packager_t::write_workbook_file(workbook_t& workbook)
{
  const std::string xml_data = workbook.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/workbook.xml");
}

void packager_t::write_worksheet_files(workbook_t& workbook)
{
  // Use ref to modify worksheet (add relations in external_hyperlinks_)
  for(size_t index = 1; auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      auto& worksheet = std::get<worksheet_t>(sheet);

      const std::string xml_data = worksheet.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/worksheets/sheet{}.xml", index));
      index++;
    }
  }
}

void packager_t::write_chartsheet_files(workbook_t& workbook)
{
  for(size_t index = 1; auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      auto& chartsheet = std::get<chartsheet_t>(sheet);

      const std::string xml_data = chartsheet.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/chartsheets/sheet{}.xml", index));
      index++;
    }
  }
}

void packager_t::write_image_files(const workbook_t& workbook)
{
  for(size_t index = 1; auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      auto& worksheet = std::get<worksheet_t>(sheet);

      if(!worksheet.image_props_.empty() || !worksheet.embedded_image_props_.empty())
      {
        for(const auto& object_props: worksheet.embedded_image_props_)
        {
          if(!object_props.is_duplicate_)
          {
            if(object_props.image_buffer_.empty())
            {
              // Read image.
              std::ifstream image_stream(object_props.filename_, std::ios::binary);
              std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
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
              std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
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
}

/// STATIC lxw_error _add_vba_project(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_error err;
///   FILE *image_stream;

///   if (!workbook->vba_project)
///     return LXW_NO_ERROR;

/* Check that the image file exists and can be opened. */
///   image_stream = lxw_fopen(workbook->vba_project, "rb");
///   if (!image_stream) {
///     LXW_WARN_FORMAT1("Error adding vbaProject.bin to xlsx file: "
///                      "file doesn't exist or can't be opened: %s.",
///                      workbook->vba_project);
///     return LXW_ERROR_CREATING_TMPFILE;
///   }

///   err = _add_file_to_zip(self, image_stream, "xl/vbaProject.bin");
///   fclose(image_stream);
///   RETURN_ON_ERROR(err);

///   return LXW_NO_ERROR;
/// }

/// STATIC lxw_error _add_vba_project_signature(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_error err;
///   FILE *image_stream;

///   if (!workbook->vba_project_signature)
///     return LXW_NO_ERROR;

/* Check that the image file exists and can be opened. */
///   image_stream = lxw_fopen(workbook->vba_project_signature, "rb");
///   if (!image_stream) {
///     LXW_WARN_FORMAT1("Error adding vbaProjectSignature.bin to xlsx file: "
///                      "file doesn't exist or can't be opened: %s.",
///                      workbook->vba_project_signature);
///     return LXW_ERROR_CREATING_TMPFILE;
///   }

///   err = _add_file_to_zip(self, image_stream, "xl/vbaProjectSignature.bin");
///   fclose(image_stream);
///   RETURN_ON_ERROR(err);

///   return LXW_NO_ERROR;
/// }

void packager_t::write_chart_files(const workbook_t& workbook)
{
  for(uint32_t index = 1; const auto& chart: workbook.ordered_charts_)
  {
    const std::string xml_data = chart->assemble_xml_file();
    add_buffer_to_zip(xml_data, std::format("xl/charts/chart{}.xml", index));
    index++;
  }
}

void packager_t::write_drawing_files(const workbook_t& workbook)
{
  for(size_t index = 1; auto& sheet: workbook.sheets_)
  {
    auto& worksheet = std::holds_alternative<worksheet_t>(sheet) ? std::get<worksheet_t>(sheet)
                                                                 : std::get<chartsheet_t>(sheet).worksheet_;
    if(worksheet.drawing_)
    {
      const std::string xml_data = worksheet.drawing_->assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/drawing{}.xml", index));
      index++;
    }
  }
}

uint32_t packager_t::get_drawing_count(const workbook_t& workbook) const
{
  uint32_t drawing_count = 0;

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      const auto& ws = std::get<chartsheet_t>(sheet).worksheet_;
      if(ws.drawing_)
      {
        drawing_count++;
      }
    }

    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);
      if(ws.drawing_)
      {
        drawing_count++;
      }
    }
  }

  return drawing_count;
}

/// STATIC lxw_error _write_table_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   lxw_table *table;
///   lxw_table_obj *table_obj;
///   lxw_error err;

///   char filename[LXW_FILENAME_LENGTH] = { 0 };
///   char *buffer = NULL;
///   size_t buffer_size = 0;
///   uint32_t index = 1;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       continue;
///     else
///       worksheet = sheet->u.worksheet;

///     if (STAILQ_EMPTY(worksheet->table_objs))
///       continue;

///     STAILQ_FOREACH(table_obj, worksheet->table_objs, list_pointers) {
///       lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                    "xl/tables/table%d.xml", index++);

///       table = lxw_table_new();
///       if (!table) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         RETURN_ON_ERROR(err);
///       }

///       table->file = lxw_get_filehandle(&buffer, &buffer_size,
///                                        self->tmpdir);
///       if (!table->file) {
///         lxw_table_free(table);
///         return LXW_ERROR_CREATING_TMPFILE;
///       }

///       table->table_obj = table_obj;

///       lxw_table_assemble_xml_file(table);

///       err = _add_to_zip(self, table->file, &buffer, &buffer_size,
///                         filename);
///       fclose(table->file);
///       free(buffer);
///       lxw_table_free(table);
///       RETURN_ON_ERROR(err);
///     }
///   }

///   return LXW_NO_ERROR;
/// }

/// uint32_t _get_table_count(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   uint32_t table_count = 0;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       worksheet = sheet->u.chartsheet->worksheet;
///     else
///       worksheet = sheet->u.worksheet;

///     table_count += worksheet->table_count;
///   }

///   return table_count;
/// }

void packager_t::write_vml_drawing_rels_file(const worksheet_t& worksheet, uint32_t index)
{
  relationships_t relationships;

  for(const auto& [type, target, target_mode]: worksheet.vml_drawing_links_)
  {
    relationships.add_worksheet_relationship(type, target, target_mode);
  }

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, std::format("xl/drawings/_rels/vmlDrawing{}.vml.rels", index));
}

void packager_t::write_vml_files(const workbook_t& workbook)
{
  uint32_t index = 1;

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);

      if(!ws.has_vml_ && !ws.has_header_vml_)
      {
        continue;
      }

      if(ws.has_vml_)
      {
        vml_t vml(ws.vml_data_id_str_, ws.comment_objs_, ws.vml_shape_id_, ws.comment_display_default_);

        ///  vml->button_objs = worksheet->button_objs;

        const std::string xml_data = vml.assemble_xml_file();
        add_buffer_to_zip(xml_data, std::format("xl/drawings/vmlDrawing{}.vml", index));
        index++;
      }

      if(ws.has_header_vml_)
      {
        write_vml_drawing_rels_file(ws, index);
        vml_t vml(ws.vml_header_id_str_, ws.header_image_objs_, ws.vml_header_id_ * 1024);
        const std::string xml_data = vml.assemble_xml_file();
        add_buffer_to_zip(xml_data, std::format("xl/drawings/vmlDrawing{}.vml", index));
        index++;
      }
    }
  }
}

void packager_t::write_comment_files(const workbook_t& workbook)
{
  uint32_t index = 1;

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);

      if(!ws.has_comments_)
      {
        continue;
      }

      comment_t comment(ws.comment_objs_, ws.comment_author_);
      const std::string xml_data = comment.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/comments{}.xml", index));
      index++;
    }
  }
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

void packager_t::write_app_file(const workbook_t& workbook)
{
  app_t app;

  if(workbook.num_worksheets_ != 0)
  {
    app.add_heading_pair("Worksheets", std::to_string(workbook.num_worksheets_));
  }

  if(workbook.num_chartsheets_ != 0)
  {
    app.add_heading_pair("Charts", std::to_string(workbook.num_chartsheets_));
  }

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);
      app.add_part_name(ws.get_sheet_name());
    }
    else if(std::holds_alternative<chartsheet_t>(sheet))
    {
      const auto& cs = std::get<chartsheet_t>(sheet);
      app.add_part_name(cs.get_sheet_name());
    }
  }

  // Add the Named Ranges parts.
  uint32_t named_range_count = 0;
  for(const auto& defined_name: workbook.defined_names_)
  {
    bool has_range  = defined_name.formula_.find('!') != std::string::npos;
    bool autofilter = defined_name.app_name_.find("_FilterDatabase") != std::string::npos;

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

  ///     app->doc_security = workbook->read_only;

  const std::string xml_data = app.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/app.xml");
}

void packager_t::write_core_file(const workbook_t& workbook)
{
  const core_t core(workbook.properties_);

  const std::string xml_data = core.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/core.xml");
}

void packager_t::write_metadata_file(const workbook_t& workbook)
{
  if(workbook.has_metadata_)
  {
    metadata_t metadata{workbook.has_dynamic_functions_, workbook.has_embedded_images_, workbook.num_embedded_images_};

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

  rich_value_t rich_value;
  const std::string xml_data = rich_value.assemble_xml_file(workbook);
  add_buffer_to_zip(xml_data, "xl/richData/rdrichvalue.xml");
}

void packager_t::write_rich_value_rel_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    rich_value_rel_t rich_value_rel{workbook.num_embedded_images_};

    const std::string xml_data = rich_value_rel.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/richValueRel.xml");
  }
}

void packager_t::write_rich_value_types_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    rich_value_types_t rich_value_types;

    const std::string xml_data = rich_value_types.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/rdRichValueTypes.xml");
  }
}

void packager_t::write_rich_value_structure_file(const workbook_t& workbook)
{
  if(workbook.has_embedded_images_)
  {
    rich_value_structure_t rich_value_structure{workbook.has_embedded_image_descriptions_};
    const std::string xml_data = rich_value_structure.assemble_xml_file();
    add_buffer_to_zip(xml_data, "xl/richData/rdrichvaluestructure.xml");
  }
}

void packager_t::write_custom_file(const workbook_t& workbook)
{
  if(workbook.custom_properties_.empty())
  {
    return;
  }

  custom_t custom(workbook.custom_properties_);

  const std::string xml_data = custom.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/custom.xml");
}

void packager_t::write_theme_file()
{
  const theme_t theme;

  const std::string xml_data = theme.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/theme/theme1.xml");
}

void packager_t::write_styles_file(const workbook_t& workbook)
{
  // TODO Manage style in workbook
  style_t styles(workbook.font_count_, workbook.fill_count_, workbook.border_count_, workbook.num_format_count_,
                 workbook.has_comments_, workbook.used_xf_formats_, workbook.used_dxf_formats_);

  const std::string xml_data = styles.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/styles.xml");
}

void packager_t::write_content_types_file(const workbook_t& workbook)
{
  content_types_t content_types;
  uint32_t worksheet_index  = 1;
  uint32_t chartsheet_index = 1;
  uint32_t drawing_count    = get_drawing_count(workbook);

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
  }

  if(!workbook.vba_project_.empty())
  {
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

  for(auto sheet: workbook.sheets_)
  {
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      content_types.add_chartsheet_name(std::format("/xl/chartsheets/sheet{}.xml", chartsheet_index));
      chartsheet_index++;
    }
    else if(std::holds_alternative<worksheet_t>(sheet))
    {
      content_types.add_worksheet_name(std::format("/xl/worksheets/sheet{}.xml", worksheet_index));
      worksheet_index++;
    }
  }

  for(size_t index = 1; index <= workbook.ordered_charts_.size(); index++)
  {
    content_types.add_chart_name(std::format("/xl/charts/chart{}.xml", index));
  }

  for(uint32_t index = 1; index <= drawing_count; index++)
  {
    content_types.add_drawing_name(std::format("/xl/drawings/drawing{}.xml", index));
  }

  ///     for (index = 1; index <= table_count; index++) {
  ///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
  ///                      "/xl/tables/table%d.xml", index);
  ///         lxw_ct_add_table_name(content_types, filename);
  ///     }

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

void packager_t::write_workbook_rels_file(const workbook_t& workbook)
{
  relationships_t relationships;
  uint32_t worksheet_index  = 1;
  uint32_t chartsheet_index = 1;

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      relationships.add_document("/chartsheet", std::format("chartsheets/sheet{}.xml", chartsheet_index));
      chartsheet_index++;
    }

    if(std::holds_alternative<worksheet_t>(sheet))
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

  ///     if (workbook->vba_project)
  ///         lxw_add_ms_package_relationship(rels, "/vbaProject",
  ///                                         "vbaProject.bin");

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

void packager_t::write_worksheet_rels_file(const workbook_t& workbook)
{
  uint32_t index = 0;

  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);
      relationships_t relationships;

      index++;

      if(ws.external_hyperlinks_.empty() && ws.external_drawing_links_.empty() &&
         ///             STAILQ_EMPTY(ws.external_table_links) &&
         !ws.external_vml_header_link_.has_value() && !ws.external_vml_comment_link_.has_value() &&
         !ws.external_background_link_.has_value() && !ws.external_comment_link_.has_value())
      {
        continue;
      }

      for(const auto& [type, target, target_mode]: ws.external_hyperlinks_)
      {
        relationships.add_worksheet_relationship(type, target, target_mode);
      }

      for(const auto& [type, target, target_mode]: ws.external_drawing_links_)
      {
        relationships.add_worksheet_relationship(type, target, target_mode);
      }

      if(ws.external_vml_comment_link_.has_value())
      {
        auto comment = ws.external_vml_comment_link_.value();
        relationships.add_worksheet_relationship(std::get<0>(comment), std::get<1>(comment), std::get<2>(comment));
      }

      if(ws.external_vml_header_link_.has_value())
      {
        auto header = ws.external_vml_header_link_.value();
        relationships.add_worksheet_relationship(std::get<0>(header), std::get<1>(header), std::get<2>(header));
      }

      if(ws.external_background_link_.has_value())
      {
        auto rel = ws.external_background_link_.value();
        relationships.add_worksheet_relationship(std::get<0>(rel), std::get<1>(rel), std::get<2>(rel));
      }

      ///         STAILQ_FOREACH(rel, worksheet->external_table_links, list_pointers) {
      ///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
      ///                                            rel->target_mode);
      ///}

      if(ws.external_comment_link_.has_value())
      {
        auto comment = ws.external_comment_link_.value();
        relationships.add_worksheet_relationship(std::get<0>(comment), std::get<1>(comment), std::get<2>(comment));
      }

      const std::string xml_data = relationships.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/worksheets/_rels/sheet{}.xml.rels", index));
    }
  }
}

void packager_t::write_chartsheet_rels_file(const workbook_t& workbook)
{
  for(size_t index = 0; const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      const auto& ws = std::get<chartsheet_t>(sheet).worksheet_;
      relationships_t relationships;

      index++;

      if(ws.external_drawing_links_.empty())
      {
        continue;
      }

      for(const auto& [type, target, target_mode]: ws.external_hyperlinks_)
      {
        relationships.add_worksheet_relationship(type, target, target_mode);
      }

      for(const auto& [type, target, target_mode]: ws.external_drawing_links_)
      {
        relationships.add_worksheet_relationship(type, target, target_mode);
      }

      const std::string xml_data = relationships.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/chartsheets/_rels/sheet{}.xml.rels", index));
    }
  }
}

void packager_t::write_drawing_rels_file(const workbook_t& workbook)
{
  for(size_t index = 1; auto& sheet: workbook.sheets_)
  {
    auto& worksheet = std::holds_alternative<worksheet_t>(sheet) ? std::get<worksheet_t>(sheet)
                                                                 : std::get<chartsheet_t>(sheet).worksheet_;
    if(!worksheet.drawing_links_.empty())
    {
      relationships_t relationships;
      for(const auto& [type, target, target_mode]: worksheet.drawing_links_)
      {
        relationships.add_worksheet_relationship(type, target, target_mode);
      }

      const std::string xml_data = relationships.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/drawings/_rels/drawing{}.xml.rels", index));
      index++;
    }
  }
}

/// STATIC lxw_error
/// _write_vba_project_rels_file(lxw_packager *self)
/// {
///     lxw_relationships *rels;
///     lxw_workbook *workbook = self->workbook;
///     lxw_error err = LXW_NO_ERROR;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!workbook->vba_project_signature)
///         return LXW_NO_ERROR;

///     rels = lxw_relationships_new();
///     if (!rels) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rels->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     lxw_add_ms_package_relationship(rels, "/vbaProjectSignature",
///                                     "vbaProjectSignature.bin");

///     lxw_relationships_assemble_xml_file(rels);

///     err = _add_to_zip(self, rels->file, &buffer, &buffer_size,
///                       "xl/_rels/vbaProject.bin.rels");

///     fclose(rels->file);
///     free(buffer);

/// mem_error:
///     lxw_free_relationships(rels);

///     return err;
/// }

void packager_t::write_rich_value_rels_file(const workbook_t& workbook)
{
  if(!workbook.has_embedded_images_)
  {
    return;
  }

  relationships_t relationships;

  for(size_t index = 1; const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& worksheet = std::get<0>(sheet);
      if(!worksheet.embedded_image_props_.empty())
      {
        for(const auto& object_props: worksheet.embedded_image_props_)
        {

          if(!object_props.is_duplicate_)
          {
            relationships.add_document("/image", std::format("../media/image{}.{}", index, object_props.extension_));
            index++;
          }
        }
      }
    }
  }

  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/richData/_rels/richValueRel.xml.rels");
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

/// STATIC lxw_error
/// _add_file_to_zip(lxw_packager *self, FILE *file, const char *filename)
/// {
///     int16_t error = ZIP_OK;
///     size_t size_read;

///     error = zipOpenNewFileInZip4_64(self->zipfile,
///                                     filename,
///                                     &self->zipfile_info,
///                                     NULL, 0, NULL, 0, NULL,
///                                     Z_DEFLATED, Z_DEFAULT_COMPRESSION, 0,
///                                     -MAX_WBITS, DEF_MEM_LEVEL,
///                                     Z_DEFAULT_STRATEGY, NULL, 0, 0, 0,
///                                     self->use_zip64);

///     if (error != ZIP_OK) {
///         LXW_ERROR("Error adding member to zipfile");
///         RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
///     }

///     fflush(file);
///     rewind(file);

///     size_read = fread((void *) self->buffer, 1, self->buffer_size, file);

///     while (size_read) {
///         if (size_read < self->buffer_size) {
///             if (ferror(file)) {
///                 LXW_ERROR("Error reading member file data");
///                 RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
///             }
///         }

///         error = zipWriteInFileInZip(self->zipfile,
///                                     self->buffer, (unsigned int) size_read);

///         if (error < 0) {
///             LXW_ERROR("Error in writing member in the zipfile");
///             RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
///         }

///         size_read =
///             fread((void *) (void *) self->buffer, 1, self->buffer_size, file);
///     }

///     error = zipCloseFileInZip(self->zipfile);
///     if (error != ZIP_OK) {
///         LXW_ERROR("Error in closing member in the zipfile");
///         RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
///     }

///     return LXW_NO_ERROR;
/// }

void packager_t::add_buffer_to_zip(std::string_view buffer, const std::string& filename)
{
  if(zipOpenNewFileInZip4_64(zipfile_, filename.c_str(), &zip_fileinfo_, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED,
                             Z_DEFAULT_COMPRESSION, 0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, nullptr, 0, 0, 0,
                             0 /*self->use_zip64*/) != ZIP_OK)
  {
    throw xwpp_exception_t(std::format("packager_t::add_buffer_to_zip(): error adding '{}' to zipfile", filename));
  }

  if(zipWriteInFileInZip(zipfile_, buffer.data(), buffer.size()) < 0)
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

void packager_t::add_buffer_to_zip(std::vector<unsigned char> buffer, const std::string& filename)
{
  if(zipOpenNewFileInZip4_64(zipfile_, filename.c_str(), &zip_fileinfo_, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED,
                             Z_DEFAULT_COMPRESSION, 0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, nullptr, 0, 0, 0,
                             0 /*self->use_zip64*/) != ZIP_OK)
  {
    throw xwpp_exception_t(
        std::format("packager_t::add_buffer_to_zip(): error adding member '{}' to zipfile", filename));
  }

  if(zipWriteInFileInZip(zipfile_, buffer.data(), buffer.size()) < 0)
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

/// STATIC lxw_error
/// _add_to_zip(lxw_packager *self, FILE *file, char **buffer,
///             size_t *buffer_size, const char *filename)
/// {
/* Flush to ensure buffer is updated when using a memory-backed file. */
///     fflush(file);
///     return *buffer ?
///         _add_buffer_to_zip(self, *buffer, *buffer_size, filename) :
///         _add_file_to_zip(self, file, filename);
/// }

// TODO const remove to allow prepare, to be refactored
void packager_t::create_package(workbook_t& workbook)
{
  /// TODO Use it as local variable (not data member) and use unique_ptr
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
  ///     error = _write_table_files(self);
  write_shared_strings_file(workbook);
  write_custom_file(workbook);
  write_theme_file();
  write_styles_file(workbook);
  write_worksheet_rels_file(workbook);
  write_chartsheet_rels_file(workbook);
  write_drawing_rels_file(workbook);
  write_image_files(workbook);
  ///     error = _add_vba_project(self);
  ///     error = _add_vba_project_signature(self);
  ///     error = _write_vba_project_rels_file(self);
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
  ///     if (zip_error) {
  ///         RETURN_ON_ZIP_ERROR(zip_error, LXW_ERROR_ZIP_CLOSE);
  ///     }
}

}
