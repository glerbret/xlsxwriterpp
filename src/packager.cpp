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
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

// TODO XML are built in-memory. Should use temporary files to improve memory usage

#include "xwpp/packager.h"

#include "xwpp/app.h"
#include "xwpp/content_types.h"
#include "xwpp/core.h"
#include "xwpp/custom.h"
#include "xwpp/exception.h"
#include "xwpp/relationships.h"
#include "xwpp/shared_strings.h"
#include "xwpp/styles.h"
#include "xwpp/theme.h"
#include "xwpp/xmlwriter.h"

#include <format>
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

/// STATIC lxw_error _write_vml_drawing_rels_file(lxw_packager *self,
///                                               lxw_worksheet *worksheet,
///                                               uint32_t index);

/// STATIC voidpf ZCALLBACK _fopen_memstream(voidpf opaque, const char *filename, int mode)
/// {
///   lxw_packager *packager = (lxw_packager *) opaque;
///   (void) filename;
///   (void) mode;
///   return lxw_get_filehandle(&packager->output_buffer,
///                             &packager->output_buffer_size,
///                             packager->tmpdir);
/// }

/// STATIC int ZCALLBACK _fclose_memstream(voidpf opaque, voidpf stream)
/// {
///   lxw_packager *packager = (lxw_packager *) opaque;
///   FILE *file = (FILE *) stream;
///   long size;

/* Ensure memstream buffer is updated */
///   if (fflush(file))
///     goto mem_error;

/* If the memstream is backed by a temporary file, no buffer is created,
   so create it manually. */
///   if (!packager->output_buffer) {
///     if (fseek(file, 0L, SEEK_END))
///       goto mem_error;

///     size = ftell(file);
///     if (size == -1)
///       goto mem_error;

///     packager->output_buffer = malloc(size);
///     GOTO_LABEL_ON_MEM_ERROR(packager->output_buffer, mem_error);

///     rewind(file);
///     if (fread((void *) packager->output_buffer, size, 1, file) < 1)
///       goto mem_error;

///     packager->output_buffer_size = size;
///   }

///   return fclose(file);

/// mem_error:
///   fclose(file);
///   return EOF;
/// }

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

void packager_t::write_workbook_file(const workbook_t& workbook)
{
  const std::string xml_data = workbook.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/workbook.xml");
}

void packager_t::write_worksheet_files(const workbook_t& workbook)
{
  /// lxw_workbook *workbook = self->workbook;
  /// lxw_sheet *sheet;
  /// lxw_worksheet *worksheet;
  /// char sheetname[LXW_FILENAME_LENGTH] = { 0 };
  /// char *buffer = NULL;
  /// size_t buffer_size = 0;
  /// uint32_t index = 1;
  /// lxw_error err;

  for(size_t index = 1; auto sheet: workbook.sheets_)
  {
    /// if (sheet->is_chartsheet)
    ///   continue;
    /// else
    {
      auto worksheet = std::get<0>(sheet);

      /// lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
      ///              "xl/worksheets/sheet%d.xml", index++);

      /// if (worksheet->optimize_row)
      ///   lxw_worksheet_write_single_row(worksheet);

      /// worksheet->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
      /// if (!worksheet->file)
      ///   return LXW_ERROR_CREATING_TMPFILE;

      const std::string xml_data = worksheet.assemble_xml_file();
      add_buffer_to_zip(xml_data, std::format("xl/worksheets/sheet{}.xml", index));
      index++;

      /// err = _add_to_zip(self, worksheet->file, &buffer,
      ///                   &buffer_size, sheetname);
    }
  }
}

/// STATIC lxw_error _write_chartsheet_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_chartsheet *chartsheet;
///   char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///   char *buffer = NULL;
///   size_t buffer_size = 0;
///   uint32_t index = 1;
///   lxw_error err;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       chartsheet = sheet->u.chartsheet;
///     else
///       continue;

///     lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,"xl/chartsheets/sheet%d.xml", index++);

///     chartsheet->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!chartsheet->file)
///       return LXW_ERROR_CREATING_TMPFILE;

///     lxw_chartsheet_assemble_xml_file(chartsheet);

///     err = _add_to_zip(self, chartsheet->file, &buffer, &buffer_size, sheetname);
///     fclose(chartsheet->file);
///     free(buffer);
///     RETURN_ON_ERROR(err);
///   }

///   return LXW_NO_ERROR;
/// }

/// STATIC lxw_error _write_image_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   lxw_object_properties *object_props;
///   lxw_error err;
///   FILE *image_stream;

///   char filename[LXW_FILENAME_LENGTH] = { 0 };
///   uint32_t index = 1;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       continue;
///     else
///       worksheet = sheet->u.worksheet;

///     if (STAILQ_EMPTY(worksheet->image_props)
///           && STAILQ_EMPTY(worksheet->embedded_image_props))
///       continue;

///     STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
///                    list_pointers) {

///       if (object_props->is_duplicate)
///        continue;

///       lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                    "xl/media/image%d.%s", index++,
///                     object_props->extension);

///       if (!object_props->is_image_buffer) {
/* Check that the image file exists and can be opened. */
///         image_stream = lxw_fopen(object_props->filename, "rb");
///         if (!image_stream) {
///           LXW_WARN_FORMAT1("Error adding image to xlsx file: file "
///                            "doesn't exist or can't be opened: %s.",
///                            object_props->filename);
///           return LXW_ERROR_CREATING_TMPFILE;
///         }

///         err = _add_file_to_zip(self, image_stream, filename);
///         fclose(image_stream);
///       }
///       else {
///         err = _add_buffer_to_zip(self,
///                                  object_props->image_buffer,
///                                  object_props->image_buffer_size,
///                                  filename);
///       }

///       RETURN_ON_ERROR(err);
///     }

///     STAILQ_FOREACH(object_props, worksheet->image_props, list_pointers) {
///       if (object_props->is_duplicate)
///         continue;

///       lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                    "xl/media/image%d.%s", index++,
///                    object_props->extension);

///       if (!object_props->is_image_buffer) {
/* Check that the image file exists and can be opened. */
///         image_stream = lxw_fopen(object_props->filename, "rb");
///         if (!image_stream) {
///           LXW_WARN_FORMAT1("Error adding image to xlsx file: file "
///                            "doesn't exist or can't be opened: %s.",
///                            object_props->filename);
///           return LXW_ERROR_CREATING_TMPFILE;
///         }

///         err = _add_file_to_zip(self, image_stream, filename);
///         fclose(image_stream);
///       }
///       else {
///         err = _add_buffer_to_zip(self,
///                                  object_props->image_buffer,
///                                  object_props->image_buffer_size,
///                                  filename);
///       }

///       RETURN_ON_ERROR(err);
///     }
///   }

///   return LXW_NO_ERROR;
/// }

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

/// STATIC lxw_error _write_chart_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_chart *chart;
///   char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///   char *buffer = NULL;
///   size_t buffer_size = 0;
///   uint32_t index = 1;
///   lxw_error err;

///   STAILQ_FOREACH(chart, workbook->ordered_charts, ordered_list_pointers) {
///     lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                  "xl/charts/chart%d.xml", index++);

///     chart->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!chart->file)
///       return LXW_ERROR_CREATING_TMPFILE;

///     lxw_chart_assemble_xml_file(chart);

///     err = _add_to_zip(self, chart->file, &buffer, &buffer_size, sheetname);
///     fclose(chart->file);
///     free(buffer);
///     RETURN_ON_ERROR(err);
///   }

///   return LXW_NO_ERROR;
/// }

/// uint32_t _get_chart_count(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_chart *chart;
///   uint32_t chart_count = 0;

///   STAILQ_FOREACH(chart, workbook->ordered_charts, ordered_list_pointers) {
///     chart_count++;
///   }

///   return chart_count;
/// }

/// STATIC lxw_error _write_drawing_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   lxw_drawing *drawing;
///   char filename[LXW_FILENAME_LENGTH] = { 0 };
///   char *buffer = NULL;
///   size_t buffer_size = 0;
///   uint32_t index = 1;
///   lxw_error err;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       worksheet = sheet->u.chartsheet->worksheet;
///     else
///       worksheet = sheet->u.worksheet;

///     drawing = worksheet->drawing;

///     if (drawing) {
///       lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                    "xl/drawings/drawing%d.xml", index++);

///       drawing->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);

///       if (!drawing->file)
///         return LXW_ERROR_CREATING_TMPFILE;

///       lxw_drawing_assemble_xml_file(drawing);

///       err = _add_to_zip(self, drawing->file, &buffer, &buffer_size, filename);
///       fclose(drawing->file);
///       free(buffer);
///       RETURN_ON_ERROR(err);
///     }
///   }

///   return LXW_NO_ERROR;
/// }

/// uint32_t _get_drawing_count(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   lxw_drawing *drawing;
///   uint32_t drawing_count = 0;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       worksheet = sheet->u.chartsheet->worksheet;
///     else
///       worksheet = sheet->u.worksheet;

///     drawing = worksheet->drawing;

///     if (drawing)
///       drawing_count++;
///   }

///   return drawing_count;
/// }

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

/// STATIC lxw_error _write_vml_files(lxw_packager *self)
/// {
///   lxw_workbook *workbook = self->workbook;
///   lxw_sheet *sheet;
///   lxw_worksheet *worksheet;
///   lxw_vml *vml;
///   char filename[LXW_FILENAME_LENGTH] = { 0 };
///   char *buffer = NULL;
///   size_t buffer_size = 0;
///   uint32_t index = 1;
///   lxw_error err;

///   STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///     if (sheet->is_chartsheet)
///       continue;
///     else
///       worksheet = sheet->u.worksheet;

///     if (!worksheet->has_vml && !worksheet->has_header_vml)
///       continue;

///     if (worksheet->has_vml) {
///       vml = lxw_vml_new();
///       if (!vml)
///         return LXW_ERROR_MEMORY_MALLOC_FAILED;

///       lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                    "xl/drawings/vmlDrawing%d.vml", index++);

///       vml->file = lxw_get_filehandle(&buffer, &buffer_size,
///                                      self->tmpdir);
///       if (!vml->file) {
///         lxw_vml_free(vml);
///         return LXW_ERROR_CREATING_TMPFILE;
///       }

///       vml->comment_objs = worksheet->comment_objs;
///       vml->button_objs = worksheet->button_objs;
///       vml->vml_shape_id = worksheet->vml_shape_id;
///       vml->comment_display_default = worksheet->comment_display_default;

///       if (worksheet->vml_data_id_str) {
///         vml->vml_data_id_str = worksheet->vml_data_id_str;
///       }
///       else {
///         fclose(vml->file);
///         free(buffer);
///         lxw_vml_free(vml);
///         return LXW_ERROR_MEMORY_MALLOC_FAILED;
///       }

///       lxw_vml_assemble_xml_file(vml);

///       err = _add_to_zip(self, vml->file, &buffer, &buffer_size,
///                         filename);

///       fclose(vml->file);
///       free(buffer);
///       lxw_vml_free(vml);

///       RETURN_ON_ERROR(err);
///     }

///     if (worksheet->has_header_vml) {

///     err = _write_vml_drawing_rels_file(self, worksheet, index);
///             RETURN_ON_ERROR(err);

///             vml = lxw_vml_new();
///             if (!vml)
///                 return LXW_ERROR_MEMORY_MALLOC_FAILED;

///             lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                          "xl/drawings/vmlDrawing%d.vml", index++);

///             vml->file = lxw_get_filehandle(&buffer, &buffer_size,
///                                            self->tmpdir);
///             if (!vml->file) {
///                 lxw_vml_free(vml);
///                 return LXW_ERROR_CREATING_TMPFILE;
///             }

///             vml->image_objs = worksheet->header_image_objs;
///             vml->vml_shape_id = worksheet->vml_header_id * 1024;

///             if (worksheet->vml_header_id_str) {
///                 vml->vml_data_id_str = worksheet->vml_header_id_str;
///             }
///             else {
///                 fclose(vml->file);
///                 free(buffer);
///                 lxw_vml_free(vml);
///                 return LXW_ERROR_MEMORY_MALLOC_FAILED;
///             }

///             lxw_vml_assemble_xml_file(vml);

///             err = _add_to_zip(self, vml->file, &buffer, &buffer_size,
///                               filename);

///             fclose(vml->file);
///             free(buffer);
///             lxw_vml_free(vml);

///             RETURN_ON_ERROR(err);
///         }
///     }

///     return LXW_NO_ERROR;
/// }

/// STATIC lxw_error
/// _write_comment_files(lxw_packager *self)
/// {
///     lxw_workbook *workbook = self->workbook;
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     lxw_comment *comment;
///     char filename[LXW_FILENAME_LENGTH] = { 0 };
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     uint32_t index = 1;
///     lxw_error err;

///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///         if (sheet->is_chartsheet)
///             continue;
///         else
///             worksheet = sheet->u.worksheet;

///         if (!worksheet->has_comments)
///             continue;

///         comment = lxw_comment_new();
///         if (!comment)
///             return LXW_ERROR_MEMORY_MALLOC_FAILED;

///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                      "xl/comments%d.xml", index++);

///         comment->file = lxw_get_filehandle(&buffer, &buffer_size,
///                                            self->tmpdir);
///         if (!comment->file) {
///             lxw_comment_free(comment);
///             return LXW_ERROR_CREATING_TMPFILE;
///         }

///         comment->comment_objs = worksheet->comment_objs;
///         comment->comment_author = worksheet->comment_author;

///         lxw_comment_assemble_xml_file(comment);

///         err = _add_to_zip(self, comment->file, &buffer, &buffer_size,
///                           filename);

///         fclose(comment->file);
///         free(buffer);
///         lxw_comment_free(comment);

///         RETURN_ON_ERROR(err);
///     }

///     return LXW_NO_ERROR;
/// }

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
  ///     lxw_workbook *workbook = self->workbook;
  ///     lxw_sheet *sheet;
  ///     lxw_worksheet *worksheet;
  ///     lxw_chartsheet *chartsheet;
  ///     lxw_defined_name *defined_name;
  ///     lxw_app *app;
  ///     char *buffer = NULL;
  ///     size_t buffer_size = 0;
  ///     uint32_t named_range_count = 0;
  ///     char *autofilter;
  ///     char *has_range;
  ///     char number[LXW_ATTR_32] = { 0 };
  ///     lxw_error err = LXW_NO_ERROR;

  if(workbook.num_worksheets_ != 0)
  {
    app.add_heading_pair("Worksheets", std::to_string(workbook.num_worksheets_));
  }

  ///     if (self->workbook->num_chartsheets) {
  ///         lxw_snprintf(number, LXW_ATTR_32, "%d",
  ///                      self->workbook->num_chartsheets);
  ///         lxw_app_add_heading_pair(app, "Charts", number);
  ///     }

  for(const auto& sheet: workbook.sheets_)
  {
    // TODO Add sheet type verification and support charsheet
    const auto& ws = std::get<worksheet_t>(sheet);
    app.add_part_name(ws.get_sheet_name());
  }
  ///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet) {
  ///             chartsheet = sheet->u.chartsheet;
  ///             lxw_app_add_part_name(app, chartsheet->name);
  ///         }
  ///     }

  /* Add the Named Ranges parts. */
  ///     TAILQ_FOREACH(defined_name, workbook->defined_names, list_pointers) {

  ///         has_range = strchr(defined_name->formula, '!');
  ///         autofilter = strstr(defined_name->app_name, "_FilterDatabase");

  /* Only store defined names with ranges (except for autofilters). */
  ///         if (has_range && !autofilter) {
  ///             lxw_app_add_part_name(app, defined_name->app_name);
  ///             named_range_count++;
  ///         }
  ///     }

  /* Add the Named Range heading pairs. */
  ///     if (named_range_count) {
  ///         lxw_snprintf(number, LXW_ATTR_32, "%d", named_range_count);
  ///         lxw_app_add_heading_pair(app, "Named Ranges", number);
  ///     }

  // Set the app/doc properties.
  app.set_properties(workbook.properties_);

  ///     app->doc_security = workbook->read_only;

  const std::string xml_data = app.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/app.xml");
}

void packager_t::write_core_file(const workbook_t& workbook)
{
  core_t core(workbook.properties_);

  const std::string xml_data = core.assemble_xml_file();
  add_buffer_to_zip(xml_data, "docProps/core.xml");
}

/// STATIC lxw_error
/// _write_metadata_file(lxw_packager *self)
/// {
///     lxw_error err = LXW_NO_ERROR;
///     lxw_metadata *metadata;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!self->workbook->has_metadata)
///         return LXW_NO_ERROR;

///     metadata = lxw_metadata_new();

///     if (!metadata) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     metadata->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!metadata->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     metadata->has_embedded_images = self->workbook->has_embedded_images;
///     metadata->num_embedded_images = self->workbook->num_embedded_images;
///     metadata->has_dynamic_functions = self->workbook->has_dynamic_functions;

///     lxw_metadata_assemble_xml_file(metadata);

///     err = _add_to_zip(self, metadata->file, &buffer, &buffer_size,
///                       "xl/metadata.xml");

///     fclose(metadata->file);
///     free(buffer);

/// mem_error:
///     lxw_metadata_free(metadata);

///     return err;
/// }

/// STATIC lxw_error
/// _write_rich_value_file(lxw_packager *self)
/// {
///     lxw_error err = LXW_NO_ERROR;
///     lxw_rich_value *rich_value;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!self->workbook->has_embedded_images)
///         return LXW_NO_ERROR;

///     rich_value = lxw_rich_value_new();
///     if (!rich_value) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     rich_value->workbook = self->workbook;

///     rich_value->file =
///         lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rich_value->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     lxw_rich_value_assemble_xml_file(rich_value);

///     err = _add_to_zip(self, rich_value->file, &buffer, &buffer_size,
///                       "xl/richData/rdrichvalue.xml");

///     fclose(rich_value->file);
///     free(buffer);

/// mem_error:
///     lxw_rich_value_free(rich_value);

///     return err;
/// }

/// STATIC lxw_error
/// _write_rich_value_rel_file(lxw_packager *self)
/// {
///     lxw_error err = LXW_NO_ERROR;
///     lxw_rich_value_rel *rich_value_rel;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!self->workbook->has_embedded_images)
///         return LXW_NO_ERROR;

///     rich_value_rel = lxw_rich_value_rel_new();
///     if (!rich_value_rel) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     rich_value_rel->num_embedded_images = self->workbook->num_embedded_images;

///     rich_value_rel->file =
///         lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rich_value_rel->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     lxw_rich_value_rel_assemble_xml_file(rich_value_rel);

///     err = _add_to_zip(self, rich_value_rel->file, &buffer, &buffer_size,
///                       "xl/richData/richValueRel.xml");

///     fclose(rich_value_rel->file);
///     free(buffer);

/// mem_error:
///     lxw_rich_value_rel_free(rich_value_rel);

///     return err;
/// }

/// STATIC lxw_error
/// _write_rich_value_types_file(lxw_packager *self)
/// {
///     lxw_error err = LXW_NO_ERROR;
///     lxw_rich_value_types *rich_value_types;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!self->workbook->has_embedded_images)
///         return LXW_NO_ERROR;

///     rich_value_types = lxw_rich_value_types_new();
///     if (!rich_value_types) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     rich_value_types->file =
///         lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rich_value_types->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     lxw_rich_value_types_assemble_xml_file(rich_value_types);

///     err = _add_to_zip(self, rich_value_types->file, &buffer, &buffer_size,
///                       "xl/richData/rdRichValueTypes.xml");

///     fclose(rich_value_types->file);
///     free(buffer);

/// mem_error:
///     lxw_rich_value_types_free(rich_value_types);

///     return err;
/// }

/// STATIC lxw_error
/// _write_rich_value_structure_file(lxw_packager *self)
/// {
///     lxw_error err = LXW_NO_ERROR;
///     lxw_rich_value_structure *rich_value_structure;
///     char *buffer = NULL;
///     size_t buffer_size = 0;

///     if (!self->workbook->has_embedded_images)
///         return LXW_NO_ERROR;

///     rich_value_structure = lxw_rich_value_structure_new();
///     if (!rich_value_structure) {
///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///         goto mem_error;
///     }

///     rich_value_structure->has_embedded_image_descriptions =
///         self->workbook->has_embedded_image_descriptions;

///     rich_value_structure->file =
///         lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rich_value_structure->file) {
///         err = LXW_ERROR_CREATING_TMPFILE;
///         goto mem_error;
///     }

///     lxw_rich_value_structure_assemble_xml_file(rich_value_structure);

///     err = _add_to_zip(self, rich_value_structure->file, &buffer, &buffer_size,
///                       "xl/richData/rdrichvaluestructure.xml");

///     fclose(rich_value_structure->file);
///     free(buffer);

/// mem_error:
///     lxw_rich_value_structure_free(rich_value_structure);

///     return err;
/// }

void packager_t::write_custom_file(const workbook_t& workbook)
{
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

void packager_t::write_styles_file(const workbook_t& /*workbook*/)
{
  // TODO Manage style in workbook
  const styles_t styles;
  ///     lxw_hash_element *hash_element;

  /* Copy the unique and in-use formats from the workbook to the styles
   * xf_format list. */
  ///     LXW_FOREACH_ORDERED(hash_element, self->workbook->used_xf_formats) {
  ///         lxw_format *workbook_format = (lxw_format *) hash_element->value;
  ///         lxw_format *style_format = lxw_format_new();

  ///         if (!style_format) {
  ///             err = LXW_ERROR_MEMORY_MALLOC_FAILED;
  ///             goto mem_error;
  ///         }

  ///         memcpy(style_format, workbook_format, sizeof(lxw_format));
  ///         STAILQ_INSERT_TAIL(styles->xf_formats, style_format, list_pointers);
  ///     }

  /* Copy the unique and in-use dxf formats from the workbook to the styles
   * dxf_format list. */
  ///     LXW_FOREACH_ORDERED(hash_element, self->workbook->used_dxf_formats) {
  ///         lxw_format *workbook_format = (lxw_format *) hash_element->value;
  ///         lxw_format *style_format = lxw_format_new();

  ///         if (!style_format) {
  ///             err = LXW_ERROR_MEMORY_MALLOC_FAILED;
  ///             goto mem_error;
  ///         }

  ///         memcpy(style_format, workbook_format, sizeof(lxw_format));
  ///         STAILQ_INSERT_TAIL(styles->dxf_formats, style_format, list_pointers);
  ///     }

  ///     styles->font_count = self->workbook->font_count;
  ///     styles->border_count = self->workbook->border_count;
  ///     styles->fill_count = self->workbook->fill_count;
  ///     styles->num_format_count = self->workbook->num_format_count;
  ///     styles->xf_count = self->workbook->used_xf_formats->unique_count;
  ///     styles->dxf_count = self->workbook->used_dxf_formats->unique_count;
  ///     styles->has_comments = self->workbook->has_comments;

  ///     styles->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
  ///     if (!styles->file) {
  ///         err = LXW_ERROR_CREATING_TMPFILE;
  ///         goto mem_error;
  ///     }

  const std::string xml_data = styles.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/styles.xml");
}

void packager_t::write_content_types_file(const workbook_t& workbook)
{
  content_types_t content_types;
  ///     lxw_content_types *content_types = lxw_content_types_new();
  ///     char *buffer = NULL;
  ///     size_t buffer_size = 0;
  ///     lxw_workbook *workbook = self->workbook;
  ///     lxw_sheet *sheet;
  ///     char filename[LXW_MAX_ATTRIBUTE_LENGTH] = { 0 };
  ///     uint32_t index = 1;
  uint32_t worksheet_index = 1;
  ///     uint32_t chartsheet_index = 1;
  ///     uint32_t drawing_count = _get_drawing_count(self);
  ///     uint32_t chart_count = _get_chart_count(self);
  ///     uint32_t table_count = _get_table_count(self);
  ///     lxw_error err = LXW_NO_ERROR;

  ///     if (!content_types) {
  ///         err = LXW_ERROR_MEMORY_MALLOC_FAILED;
  ///         goto mem_error;
  ///     }

  ///     content_types->file = lxw_get_filehandle(&buffer, &buffer_size,
  ///                                              self->tmpdir);
  ///     if (!content_types->file) {
  ///         err = LXW_ERROR_CREATING_TMPFILE;
  ///         goto mem_error;
  ///     }

  ///     if (workbook->has_png)
  ///         lxw_ct_add_default(content_types, "png", "image/png");

  ///     if (workbook->has_jpeg)
  ///         lxw_ct_add_default(content_types, "jpeg", "image/jpeg");

  ///     if (workbook->has_bmp)
  ///         lxw_ct_add_default(content_types, "bmp", "image/bmp");

  ///     if (workbook->has_gif)
  ///         lxw_ct_add_default(content_types, "gif", "image/gif");

  ///     if (workbook->vba_project)
  ///         lxw_ct_add_default(content_types, "bin",
  ///                            "application/vnd.ms-office.vbaProject");

  ///     if (workbook->vba_project)
  ///         lxw_ct_add_override(content_types, "/xl/workbook.xml",
  ///                             LXW_APP_MSEXCEL
  ///                             "sheet.macroEnabled.main+xml");
  ///     else
  content_types.add_override("/xl/workbook.xml", content_types_t::APP_DOCUMENT + "spreadsheetml.sheet.main+xml");

  ///     if (workbook->vba_project_signature)
  ///         lxw_ct_add_override(content_types, "/xl/vbaProjectSignature.bin",
  ///                             "application/vnd.ms-office.vbaProjectSignature");

  for(auto sheet: workbook.sheets_)
  {
    ///         if (sheet->is_chartsheet) {
    ///             lxw_snprintf(filename, LXW_FILENAME_LENGTH,
    ///                          "/xl/chartsheets/sheet%d.xml",
    ///                          chartsheet_index++);
    ///             lxw_ct_add_chartsheet_name(content_types, filename);
    ///         }
    ///         else {
    content_types.add_worksheet_name(std::format("/xl/worksheets/sheet{}.xml", worksheet_index));
    worksheet_index++;
    ///         }
  }

  ///     for (index = 1; index <= chart_count; index++) {
  ///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
  ///         "/xl/charts/chart%d.xml",
  ///                      index);
  ///         lxw_ct_add_chart_name(content_types, filename);
  ///     }

  ///     for (index = 1; index <= drawing_count; index++) {
  ///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
  ///                      "/xl/drawings/drawing%d.xml", index);
  ///         lxw_ct_add_drawing_name(content_types, filename);
  ///     }

  ///     for (index = 1; index <= table_count; index++) {
  ///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
  ///                      "/xl/tables/table%d.xml", index);
  ///         lxw_ct_add_table_name(content_types, filename);
  ///     }

  ///     if (workbook->has_vml)
  ///         lxw_ct_add_vml_name(content_types);

  ///     for (index = 1; index <= workbook->comment_count; index++) {
  ///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
  ///                      "/xl/comments%d.xml", index);
  ///         lxw_ct_add_comment_name(content_types, filename);
  ///     }

  if(workbook.sst_.has_string())
  {
    content_types.add_shared_strings();
  }

  if(!workbook.custom_properties_.empty())
  {
    content_types.add_custom_properties();
  }

  ///     if (workbook->has_metadata)
  ///         lxw_ct_add_metadata(content_types);

  ///     if (workbook->has_embedded_images)
  ///         lxw_ct_add_rich_value(content_types);

  const std::string xml_data = content_types.assemble_xml_file();
  add_buffer_to_zip(xml_data, "[Content_Types].xml");
}

void packager_t::write_workbook_rels_file(const workbook_t& workbook)
{
  relationships_t relationships;
  ///     char *buffer = NULL;
  ///     size_t buffer_size = 0;
  ///     lxw_workbook *workbook = self->workbook;
  ///     lxw_sheet *sheet;
  ///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
  uint32_t worksheet_index = 1;
  ///     uint32_t chartsheet_index = 1;
  ///     lxw_error err = LXW_NO_ERROR;

  ///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet) {
  ///             lxw_snprintf(sheetname,
  ///                          LXW_FILENAME_LENGTH,
  ///                          "chartsheets/sheet%d.xml", chartsheet_index++);
  ///             lxw_add_document_relationship(rels, "/chartsheet", sheetname);
  ///         }
  ///         else {
  ///             lxw_snprintf(sheetname,
  ///                          LXW_FILENAME_LENGTH,
  ///                          "worksheets/sheet%d.xml", worksheet_index++);
  relationships.add_document("/worksheet", std::format("worksheets/sheet{}.xml", worksheet_index));
  worksheet_index++;
  ///         }
  ///     }

  relationships.add_document("/theme", "theme/theme1.xml");
  relationships.add_document("/styles", "styles.xml");

  if(workbook.sst_.has_string())
  {
    relationships.add_document("/sharedStrings", "sharedStrings.xml");
  }

  ///     if (workbook->vba_project)
  ///         lxw_add_ms_package_relationship(rels, "/vbaProject",
  ///                                         "vbaProject.bin");

  ///     if (workbook->has_metadata)
  ///         lxw_add_document_relationship(rels, "/sheetMetadata",
  ///         "metadata.xml");

  ///     if (workbook->has_embedded_images)
  ///         lxw_add_rich_value_relationship(rels);
  const std::string xml_data = relationships.assemble_xml_file();
  add_buffer_to_zip(xml_data, "xl/_rels/workbook.xml.rels");
}

/// STATIC lxw_error
/// _write_worksheet_rels_file(lxw_packager *self)
/// {
///     lxw_relationships *rels;
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     lxw_rel_tuple *rel;
///     lxw_workbook *workbook = self->workbook;
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///     uint32_t index = 0;
///     lxw_error err;

///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///         if (sheet->is_chartsheet)
///             continue;
///         else
///             worksheet = sheet->u.worksheet;

///         index++;

///         if (STAILQ_EMPTY(worksheet->external_hyperlinks) &&
///             STAILQ_EMPTY(worksheet->external_drawing_links) &&
///             STAILQ_EMPTY(worksheet->external_table_links) &&
///             !worksheet->external_vml_header_link &&
///             !worksheet->external_vml_comment_link &&
///             !worksheet->external_background_link &&
///             !worksheet->external_comment_link)
///             continue;

///         rels = lxw_relationships_new();

///         rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///         if (!rels->file) {
///             lxw_free_relationships(rels);
///             return LXW_ERROR_CREATING_TMPFILE;
///         }

///         STAILQ_FOREACH(rel, worksheet->external_hyperlinks, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         STAILQ_FOREACH(rel, worksheet->external_drawing_links, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         rel = worksheet->external_vml_comment_link;
///         if (rel)
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);

///         rel = worksheet->external_vml_header_link;
///         if (rel)
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);

///         rel = worksheet->external_background_link;
///         if (rel)
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);

///         STAILQ_FOREACH(rel, worksheet->external_table_links, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         rel = worksheet->external_comment_link;
///         if (rel)
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);

///         lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                      "xl/worksheets/_rels/sheet%d.xml.rels", index);

///         lxw_relationships_assemble_xml_file(rels);

///         err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

///         fclose(rels->file);
///         free(buffer);
///         lxw_free_relationships(rels);

///         RETURN_ON_ERROR(err);
///     }

///     return LXW_NO_ERROR;
/// }

/// STATIC lxw_error
/// _write_chartsheet_rels_file(lxw_packager *self)
/// {
///     lxw_relationships *rels;
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     lxw_rel_tuple *rel;
///     lxw_workbook *workbook = self->workbook;
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///     uint32_t index = 0;
///     lxw_error err;

///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///         if (sheet->is_chartsheet)
///             worksheet = sheet->u.chartsheet->worksheet;
///         else
///             continue;

///         index++;

///         if (STAILQ_EMPTY(worksheet->external_drawing_links))
///             continue;

///         rels = lxw_relationships_new();

///         rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///         if (!rels->file) {
///             lxw_free_relationships(rels);
///             return LXW_ERROR_CREATING_TMPFILE;
///         }

///         STAILQ_FOREACH(rel, worksheet->external_hyperlinks, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         STAILQ_FOREACH(rel, worksheet->external_drawing_links, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                      "xl/chartsheets/_rels/sheet%d.xml.rels", index);

///         lxw_relationships_assemble_xml_file(rels);

///         err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

///         fclose(rels->file);
///         free(buffer);
///         lxw_free_relationships(rels);

///         RETURN_ON_ERROR(err);
///     }

///     return LXW_NO_ERROR;
/// }

/// STATIC lxw_error
/// _write_drawing_rels_file(lxw_packager *self)
/// {
///     lxw_relationships *rels;
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     lxw_rel_tuple *rel;
///     lxw_workbook *workbook = self->workbook;
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///     uint32_t index = 1;
///     lxw_error err;

///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///         if (sheet->is_chartsheet)
///             worksheet = sheet->u.chartsheet->worksheet;
///         else
///             worksheet = sheet->u.worksheet;

///         if (STAILQ_EMPTY(worksheet->drawing_links))
///             continue;

///         rels = lxw_relationships_new();

///         rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///         if (!rels->file) {
///             lxw_free_relationships(rels);
///             return LXW_ERROR_CREATING_TMPFILE;
///         }

///         STAILQ_FOREACH(rel, worksheet->drawing_links, list_pointers) {
///             lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                            rel->target_mode);
///         }

///         lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                      "xl/drawings/_rels/drawing%d.xml.rels", index++);

///         lxw_relationships_assemble_xml_file(rels);

///         err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

///         fclose(rels->file);
///         free(buffer);
///         lxw_free_relationships(rels);

///         RETURN_ON_ERROR(err);
///     }

///     return LXW_NO_ERROR;
/// }

/// STATIC lxw_error
/// _write_vml_drawing_rels_file(lxw_packager *self, lxw_worksheet *worksheet,
///                              uint32_t index)
/// {
///     lxw_relationships *rels;
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     lxw_rel_tuple *rel;
///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///     lxw_error err = LXW_NO_ERROR;

///     rels = lxw_relationships_new();

///     rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rels->file) {
///         lxw_free_relationships(rels);
///         return LXW_ERROR_CREATING_TMPFILE;
///     }

///     STAILQ_FOREACH(rel, worksheet->vml_drawing_links, list_pointers) {
///         lxw_add_worksheet_relationship(rels, rel->type, rel->target,
///                                        rel->target_mode);
///     }

///     lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                  "xl/drawings/_rels/vmlDrawing%d.vml.rels", index);

///     lxw_relationships_assemble_xml_file(rels);

///     err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

///     fclose(rels->file);
///     free(buffer);
///     lxw_free_relationships(rels);

///     return err;
/// }

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

/// STATIC lxw_error
/// _write_rich_value_rels_file(lxw_packager *self)
/// {
///     lxw_workbook *workbook = self->workbook;
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     lxw_object_properties *object_props;

///     lxw_relationships *rels;
///     char *buffer = NULL;
///     size_t buffer_size = 0;
///     char sheetname[LXW_FILENAME_LENGTH] = { 0 };
///     char target[LXW_FILENAME_LENGTH] = { 0 };
///     lxw_error err = LXW_NO_ERROR;
///     uint32_t index = 1;

///     if (!workbook->has_embedded_images)
///         return LXW_NO_ERROR;

///     rels = lxw_relationships_new();

///     rels->file = lxw_get_filehandle(&buffer, &buffer_size, self->tmpdir);
///     if (!rels->file) {
///         lxw_free_relationships(rels);
///         return LXW_ERROR_CREATING_TMPFILE;
///     }

///     STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
///         if (sheet->is_chartsheet)
///             continue;
///         else
///             worksheet = sheet->u.worksheet;

///         if (STAILQ_EMPTY(worksheet->embedded_image_props))
///             continue;

///         STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
///                        list_pointers) {

///             if (object_props->is_duplicate)
///                 continue;

///             lxw_snprintf(target, LXW_FILENAME_LENGTH,
///                          "../media/image%d.%s", index++,
///                          object_props->extension);

///             lxw_add_document_relationship(rels, "/image", target);
///         }
///     }

///     lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
///                  "xl/richData/_rels/richValueRel.xml.rels");

///     lxw_relationships_assemble_xml_file(rels);

///     err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

///     fclose(rels->file);
///     free(buffer);
///     lxw_free_relationships(rels);

///     return err;
/// }

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
    throw xwpp_exception_t(std::format("Error adding member {} to zipfile", filename));
  }

  if(zipWriteInFileInZip(zipfile_, buffer.data(), buffer.size()) < 0)
  {
    throw xwpp_exception_t(std::format("Error in writing member {} to zipfile", filename));
  }

  if(zipCloseFileInZip(zipfile_) != ZIP_OK)
  {
    throw xwpp_exception_t(std::format("Error in closing member {} to zipfile", filename));
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

void packager_t::create_package(const workbook_t& workbook)
{
  ///     lxw_error error;
  ///     int8_t zip_error;
  // Create a zip container for the xlsx file.
  /// TODO Use it as local variable (not data member) and use unique_ptr
  zipfile_ = zipOpen(filename_.c_str(), 0);

  write_content_types_file(workbook);
  write_root_rels_file(workbook);
  write_workbook_rels_file(workbook);
  write_worksheet_files(workbook);
  ///     error = _write_chartsheet_files(self);
  write_workbook_file(workbook);
  ///     error = _write_chart_files(self);
  ///     error = _write_drawing_files(self);
  ///     error = _write_vml_files(self);
  ///     error = _write_comment_files(self);
  ///     error = _write_table_files(self);
  write_shared_strings_file(workbook);
  write_custom_file(workbook);
  write_theme_file();
  write_styles_file(workbook);
  ///     error = _write_worksheet_rels_file(self);
  ///     error = _write_chartsheet_rels_file(self);
  ///     error = _write_drawing_rels_file(self);
  ///     error = _write_image_files(self);
  ///     error = _add_vba_project(self);
  ///     error = _add_vba_project_signature(self);
  ///     error = _write_vba_project_rels_file(self);
  write_core_file(workbook);
  ///     error = _write_metadata_file(self);
  ///     error = _write_rich_value_file(self);
  ///     error = _write_rich_value_rel_file(self);
  ///     error = _write_rich_value_types_file(self);
  ///     error = _write_rich_value_structure_file(self);
  ///     error = _write_rich_value_rels_file(self);
  write_app_file(workbook);

  // TODO Manage errors (exception)
  /*zip_error =*/zipClose(zipfile_, nullptr);
  ///     if (zip_error) {
  ///         RETURN_ON_ZIP_ERROR(zip_error, LXW_ERROR_ZIP_CLOSE);
  ///     }
}

}
