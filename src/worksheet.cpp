/*
 * worksheet - A library for creating Excel XLSX worksheet files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

/// #ifdef USE_FMEMOPEN
/// #define _POSIX_C_SOURCE 200809L
/// #endif

#include "xwpp/worksheet.h"

#include "xwpp/drawing.h"
#include "xwpp/exception.h"
#include "xwpp/shared_strings.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <openssl/md5.h>

#include <cstdint>
#include <format>
#include <fstream>
#include <span>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#include <iostream>

using namespace std::literals;

namespace xwpp
{

namespace
{

/*
 * Calculate the "spans" attribute of the <row> tag. This is an XLSX
 * optimization and isn't strictly required. However, it makes comparing
 * files easier.
 *
 * The span is the same for each block of 16 rows.
 */
std::string calculate_spans(std::map<col_num_t, row_t>::const_iterator it,
                            std::map<col_num_t, row_t>::const_iterator end, int32_t& block_num)
{
  col_num_t span_col_min = std::numeric_limits<col_num_t>::max();
  col_num_t span_col_max = std::numeric_limits<col_num_t>::max();

  block_num = it->second.row_num_ / 16;

  for(; it != end && static_cast<int32_t>(it->second.row_num_ / 16) == block_num; it++)
  {
    if(!it->second.cells_.empty())
    {
      for(const auto& [col_num, cell]: it->second.cells_)
      {
        if(col_num < span_col_min || span_col_min == std::numeric_limits<col_num_t>::max())
        {
          span_col_min = col_num;
        }
        if(col_num > span_col_max || span_col_max == std::numeric_limits<col_num_t>::max())
        {
          span_col_max = col_num;
        }
      }
    }
  }

  return std::format("{}:{}", span_col_min + 1, span_col_max + 1);
}

}

/// #define LXW_BUFFER_SIZE                  4096
/// #define LXW_PRINT_ACROSS                 1
/// #define LXW_VALIDATION_MAX_TITLE_LENGTH  32
/// #define LXW_VALIDATION_MAX_STRING_LENGTH 255
/// #define LXW_THIS_ROW "[#This Row],"

/// STATIC int _row_cmp(row_t *row1, row_t *row2);
/// STATIC int _cell_cmp(cell_t *cell1, cell_t *cell2);
/// STATIC int _drawing_rel_id_cmp(lxw_drawing_rel_id *tuple1,
///                                lxw_drawing_rel_id *tuple2);
/// STATIC int _cond_format_hash_cmp(lxw_cond_format_hash_element *elem_1,
///                                  lxw_cond_format_hash_element *elem_2);

const row_t* worksheet_t::find_row(row_num_t row_num) const
{
  auto it = table_.rbh_root_.find(row_num);
  if(it != std::end(table_.rbh_root_))
  {
    return &it->second;
  }
  else
  {
    return nullptr;
  }
}

const cell_t* worksheet_t::find_cell_in_row(const row_t* row, col_num_t col_num) const
{
  if(!row)
  {
    return nullptr;
  }

  auto it = row->cells_.find(col_num);
  if(it != std::end(row->cells_))
  {
    return &it->second;
  }
  else
  {
    return nullptr;
  }
}

worksheet_t::worksheet_t(const worksheet_init_data_t& init_data, std::function<int32_t(format_t*)> get_xf_index)
  : get_xf_index_{get_xf_index}
  , sst_{init_data.sst_}
  , name_{init_data.name_}
  , quoted_name_{init_data.quoted_name_}
  , index_{init_data.index_}
  , default_url_format_{init_data.default_url_format_}
  , header_footer_objs_{std::nullopt, std::nullopt, std::nullopt, std::nullopt, std::nullopt, std::nullopt}
{

  /* Initialize the cached rows. */
  ///     worksheet->table->cached_row_num = LXW_ROW_MAX + 1;
  ///     worksheet->hyperlinks->cached_row_num = LXW_ROW_MAX + 1;
  ///     worksheet->comments->cached_row_num = LXW_ROW_MAX + 1;

  ///     if (init_data && init_data->optimize) {
  ///         worksheet->array = calloc(LXW_COL_MAX, sizeof(struct cell_t *));
  ///         GOTO_LABEL_ON_MEM_ERROR(worksheet->array, mem_error);
  ///     }

  ///     worksheet->optimize_row = calloc(1, sizeof(struct row_t));
  ///     GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_row, mem_error);
  ///     worksheet->optimize_row->height = LXW_DEF_ROW_HEIGHT;

  ///     if (init_data && init_data->optimize) {
  ///         FILE *tmpfile;

  ///         worksheet->optimize_buffer = NULL;
  ///         worksheet->optimize_buffer_size = 0;
  ///         tmpfile = lxw_get_filehandle(&worksheet->optimize_buffer,
  ///                                      &worksheet->optimize_buffer_size,
  ///                                      init_data->tmpdir);
  ///         if (!tmpfile) {
  ///             LXW_ERROR("Error creating tmpfile() for worksheet in "
  ///                       "'constant_memory' mode.");
  ///             goto mem_error;
  ///         }

  ///         worksheet->optimize_tmpfile = tmpfile;
  ///         GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_tmpfile, mem_error);
  ///         worksheet->file = worksheet->optimize_tmpfile;
  ///     }

  ///     if (init_data) {
  ///         worksheet->tmpdir = init_data->tmpdir;
  ///         worksheet->hidden = init_data->hidden;
  ///         worksheet->sst = init_data->sst;
  ///         worksheet->optimize = init_data->optimize;
  ///         worksheet->active_sheet = init_data->active_sheet;
  ///         worksheet->first_sheet = init_data->first_sheet;
  ///         worksheet->max_url_length = init_data->max_url_length;
  ///         worksheet->use_1904_epoch = init_data->use_1904_epoch;
  ///     }
}

namespace
{

cell_t new_string_cell(row_num_t row_num, col_num_t col_num, uint32_t string_id, const std::string& sst_string,
                       const format_t* format)
{
  cell_t cell;

  cell.row_num_    = row_num;
  cell.col_num_    = col_num;
  cell.type_       = cell_types_t::STRING_CELL;
  cell.format_     = const_cast<format_t*>(format);
  cell.data_       = string_id;
  cell.sst_string_ = sst_string;

  return cell;
}

cell_t new_number_cell(row_num_t row_num, col_num_t col_num, double value, const format_t* format)
{
  cell_t cell;

  cell.row_num_ = row_num;
  cell.col_num_ = col_num;
  cell.type_    = cell_types_t::NUMBER_CELL;
  cell.format_  = const_cast<format_t*>(format);
  cell.data_    = value;

  return cell;
}

cell_t new_hyperlink_cell(row_num_t row_num, col_num_t col_num, cell_types_t link_type, const std::string& url,
                          const std::string& str, const std::string& tooltip)
{
  cell_t cell;

  cell.row_num_    = row_num;
  cell.col_num_    = col_num;
  cell.type_       = link_type;
  cell.data_       = url;
  cell.user_data1_ = str;
  cell.user_data2_ = tooltip;

  return cell;
}

cell_t new_comment_cell(row_num_t row_num, col_num_t col_num, const vml_obj_t& comment)
{
  cell_t cell;

  cell.row_num_ = row_num;
  cell.col_num_ = col_num;
  cell.type_    = cell_types_t::COMMENT;
  cell.comment_ = comment;

  return cell;
}

cell_t new_blank_cell(row_num_t row_num, col_num_t col_num, const format_t* format)
{
  cell_t cell;

  cell.row_num_ = row_num;
  cell.col_num_ = col_num;
  cell.type_    = cell_types_t::BLANK_CELL;
  cell.format_  = const_cast<format_t*>(format);

  return cell;
}

cell_t new_error_cell(row_num_t row_num, col_num_t col_num, uint32_t value, const format_t* format)
{
  cell_t cell;

  cell.row_num_ = row_num;
  cell.col_num_ = col_num;
  cell.type_    = cell_types_t::ERROR_CELL;
  cell.format_  = const_cast<format_t*>(format);
  cell.data_    = value;

  return cell;
}

cell_t new_formula_cell(row_num_t row_num, col_num_t col_num, const std::string& formula, const format_t* format,
                        double result)
{
  cell_t cell;

  cell.row_num_        = row_num;
  cell.col_num_        = col_num;
  cell.type_           = cell_types_t::FORMULA_CELL;
  cell.format_         = const_cast<format_t*>(format);
  cell.data_           = formula;
  cell.formula_result_ = result;

  return cell;
}

cell_t new_formula_cell(row_num_t row_num, col_num_t col_num, const std::string& formula, const format_t* format,
                        const std::string& result)
{
  cell_t cell;

  cell.row_num_    = row_num;
  cell.col_num_    = col_num;
  cell.type_       = cell_types_t::FORMULA_CELL;
  cell.format_     = const_cast<format_t*>(format);
  cell.data_       = formula;
  cell.user_data2_ = result;

  return cell;
}

/*
 * This function handles the additional optional parameters to
 * worksheet_write_comment_opt() as well as calculating the comment object
 * position and vertices.
 */
void get_comment_params(vml_obj_t& comment, std::optional<comment_options_t> options)
{
  row_num_t start_row;
  col_num_t start_col;
  int32_t x_offset;
  int32_t y_offset;
  uint32_t height = 74;
  uint32_t width  = 128;
  double x_scale  = 1.0;
  double y_scale  = 1.0;
  row_num_t row   = comment.row_;
  col_num_t col   = comment.col_;

  /* Set the default start cell and offsets for the comment. These are
   * generally fixed in relation to the parent cell. However there are some
   * edge cases for cells at the, well yes, edges. */
  if(row == 0)
  {
    y_offset = 2;
  }
  else if(row == worksheet_t::ROW_MAX - 3)
  {
    y_offset = 16;
  }
  else if(row == worksheet_t::ROW_MAX - 2)
  {
    y_offset = 16;
  }
  else if(row == worksheet_t::ROW_MAX - 1)
  {
    y_offset = 14;
  }
  else
  {
    y_offset = 10;
  }

  if(col == worksheet_t::COL_MAX - 3)
  {
    x_offset = 49;
  }
  else if(col == worksheet_t::COL_MAX - 2)
  {
    x_offset = 49;
  }
  else if(col == worksheet_t::COL_MAX - 1)
  {
    x_offset = 49;
  }
  else
  {
    x_offset = 15;
  }

  if(row == 0)
  {
    start_row = 0;
  }
  else if(row == worksheet_t::ROW_MAX - 3)
  {
    start_row = worksheet_t::ROW_MAX - 7;
  }
  else if(row == worksheet_t::ROW_MAX - 2)
  {
    start_row = worksheet_t::ROW_MAX - 6;
  }
  else if(row == worksheet_t::ROW_MAX - 1)
  {
    start_row = worksheet_t::ROW_MAX - 5;
  }
  else
  {
    start_row = row - 1;
  }

  if(col == worksheet_t::COL_MAX - 3)
  {
    start_col = worksheet_t::COL_MAX - 6;
  }
  else if(col == worksheet_t::COL_MAX - 2)
  {
    start_col = worksheet_t::COL_MAX - 5;
  }
  else if(col == worksheet_t::COL_MAX - 1)
  {
    start_col = worksheet_t::COL_MAX - 4;
  }
  else
  {
    start_col = col + 1;
  }

  // Set the default font properties.
  comment.font_size_   = 8;
  comment.font_family_ = 2;

  // Set any user defined options.
  if(options)
  {
    if(options->width_ > 0.0)
    {
      width = options->width_;
    }

    if(options->height_ > 0.0)
    {
      height = options->height_;
    }

    if(options->x_scale_ > 0.0)
    {
      x_scale = options->x_scale_;
    }

    if(options->y_scale_ > 0.0)
    {
      y_scale = options->y_scale_;
    }

    if(options->x_offset_ != 0)
    {
      x_offset = options->x_offset_;
    }

    if(options->y_offset_ != 0)
    {
      y_offset = options->y_offset_;
    }

    if(options->start_row_ > 0 || options->start_col_ > 0)
    {
      start_row = options->start_row_;
      start_col = options->start_col_;
    }

    if(options->font_size_ > 0.0)
    {
      comment.font_size_ = options->font_size_;
    }

    if(options->font_family_ > 0)
    {
      comment.font_family_ = options->font_family_;
    }

    comment.visible_   = options->visible_;
    comment.color_     = options->color_;
    comment.author_    = options->author_;
    comment.font_name_ = options->font_name_;
  }

  // Scale the width/height to the default/user scale and round to the
  // nearest pixel.
  width  = (uint32_t)(0.5 + x_scale * width);
  height = (uint32_t)(0.5 + y_scale * height);

  comment.width_     = width;
  comment.height_    = height;
  comment.start_col_ = start_col;
  comment.start_row_ = start_row;
  comment.x_offset_  = x_offset;
  comment.y_offset_  = y_offset;
}

}

/// STATIC cell_t * _new_inline_string_cell(row_num_t row_num,
///                         col_num_t col_num, char *string, lxw_format *format)
/// {
///     cell_t *cell = calloc(1, sizeof(cell_t));
///     RETURN_ON_MEM_ERROR(cell, cell);

///     cell->row_num = row_num;
///     cell->col_num = col_num;
///     cell->type = INLINE_STRING_CELL;
///     cell->format = format;
///     cell->u.string = string;

///     return cell;
/// }

/// STATIC cell_t *
/// _new_inline_rich_string_cell(row_num_t row_num,
///                              col_num_t col_num, const char *string,
///                              lxw_format *format)
/// {
///     cell_t *cell = calloc(1, sizeof(cell_t));
///     RETURN_ON_MEM_ERROR(cell, cell);

///     cell->row_num = row_num;
///     cell->col_num = col_num;
///     cell->type = INLINE_RICH_STRING_CELL;
///     cell->format = format;
///     cell->u.string = string;

///     return cell;
/// }

/// STATIC cell_t *
/// _new_array_formula_cell(row_num_t row_num, col_num_t col_num, char *formula,
///                         char *range, lxw_format *format, uint8_t is_dynamic)
/// {
///     cell_t *cell = calloc(1, sizeof(cell_t));
///     RETURN_ON_MEM_ERROR(cell, cell);

///     cell->row_num = row_num;
///     cell->col_num = col_num;
///     cell->format = format;
///     cell->u.string = formula;
///     cell->user_data1 = range;

///     if (is_dynamic)
///         cell->type = DYNAMIC_ARRAY_FORMULA_CELL;
///     else
///         cell->type = ARRAY_FORMULA_CELL;

///     return cell;
/// }

/// STATIC cell_t *
/// _new_boolean_cell(row_num_t row_num, col_num_t col_num, int value,
///                   lxw_format *format)
/// {
///     cell_t *cell = calloc(1, sizeof(cell_t));
///     RETURN_ON_MEM_ERROR(cell, cell);

///     cell->row_num = row_num;
///     cell->col_num = col_num;
///     cell->type = BOOLEAN_CELL;
///     cell->format = format;
///     cell->u.number = value;

///     return cell;
/// }

row_t& table_rows_t::get_row_list(row_num_t row_num)
{
  row_t& row = rbh_root_[row_num];

  if(row.row_num_ == 0)
  {
    // New created row, initialise it
    row.row_num_ = row_num;
  }

  ///    table->cached_row = row;
  ///    table->cached_row_num = row_num;
  return row;
}

row_t& worksheet_t::get_row(row_num_t row_num)
{
  ///    if (!self->optimize) {
  return table_.get_row_list(row_num);
  ///    }
  ///    else {
  ///        if (row_num < self->optimize_row->row_num) {
  ///            return NULL;
  ///        }
  ///        else if (row_num == self->optimize_row->row_num) {
  ///            return self->optimize_row;
  ///        }
  ///        else {
  /* Flush row. */
  ///            lxw_worksheet_write_single_row(self);
  ///            row = self->optimize_row;
  ///            row->row_num = row_num;
  ///            return row;
  ///        }
  ///    }
}

void worksheet_t::insert_cell(row_num_t row_num, col_num_t col_num, const cell_t& cell)
{
  row_t& row = get_row(row_num);

  ///    if (!self->optimize) {
  row.data_changed_   = true;
  row.cells_[col_num] = cell;
  ///    }
  ///    else {
  ///        if (row) {
  ///            row->data_changed = LXW_TRUE;

  /* Overwrite an existing cell if necessary. */
  ///            if (self->array[col_num])
  ///                _free_cell(self->array[col_num]);

  ///            self->array[col_num] = cell;
  ///        }
  ///    }
}

/*
 * Insert a blank placeholder cell in the cells RB tree in the same position
 * as a comment so that the rows "spans" calculation is correct. Since the
 * blank cell doesn't have a format it is ignored when writing. If there is
 * already a cell in the required position we don't have add a new cell.
 */
void worksheet_t::insert_cell_placeholder(row_num_t row_num, col_num_t col_num)
{
  /* The spans calculation isn't required in constant_memory mode. */
  ///     if (self->optimize)
  ///         return;

  // Only add a cell if one doesn't already exist.
  row_t& row = get_row(row_num);
  if(row.cells_.find(col_num) == std::end(row.cells_))
  {
    cell_t cell         = new_blank_cell(row_num, col_num, nullptr);
    row.cells_[col_num] = cell;
  }
}

void worksheet_t::insert_hyperlink(row_num_t row_num, col_num_t col_num, const cell_t& link)
{
  row_t& row = hyperlinks_.get_row_list(row_num);

  row.cells_[col_num] = link;
}

void worksheet_t::insert_comment(row_num_t row_num, col_num_t col_num, const cell_t& link)
{
  row_t& row = comments_.get_row_list(row_num);

  row.cells_[col_num] = link;
}

/// STATIC col_num_t
/// _next_power_of_two(uint16_t col)
/// {
///     col--;
///     col |= col >> 1;
///     col |= col >> 2;
///     col |= col >> 4;
///     col |= col >> 8;
///     col++;

///     return col;
/// }

/*
 * Check that row and col are within the allowed Excel range and store max
 * and min values for use in other methods/elements.
 *
 * The ignore_row/ignore_col flags are used to indicate that we wish to
 * perform the dimension check without storing the value.
 */
void worksheet_t::check_dimensions(row_num_t row_num, col_num_t col_num, bool ignore_row, bool ignore_col)
{
  if(row_num >= ROW_MAX)
  {
    throw xwpp_out_of_range_t(std::format("Row {} out of range (max {})", row_num, ROW_MAX));
  }

  if(col_num >= COL_MAX)
  {
    throw xwpp_out_of_range_t(std::format("Col {} out of range (max {})", col_num, COL_MAX));
  }

  /* In optimization mode we don't change dimensions for rows that are */
  /* already written. */
  ///     if (!ignore_row && !ignore_col && self->optimize) {
  ///         if (row_num < self->optimize_row->row_num)
  ///             return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
  ///     }

  if(!ignore_row)
  {
    dim_rowmin_ = std::min(row_num, dim_rowmin_);
    dim_rowmax_ = std::max(row_num, dim_rowmax_);
  }

  if(!ignore_col)
  {
    dim_colmin_ = std::min(col_num, dim_colmin_);
    dim_colmax_ = std::max(col_num, dim_colmax_);
  }
}

/// STATIC int
/// _row_cmp(row_t *row1, row_t *row2)
/// {
///     if (row1->row_num > row2->row_num)
///         return 1;
///     if (row1->row_num < row2->row_num)
///         return -1;
///     return 0;
/// }

/// STATIC int
/// _cell_cmp(cell_t *cell1, cell_t *cell2)
/// {
///     if (cell1->col_num > cell2->col_num)
///         return 1;
///     if (cell1->col_num < cell2->col_num)
///         return -1;
///     return 0;
/// }

/// STATIC int _drawing_rel_id_cmp(lxw_drawing_rel_id *rel_id1, lxw_drawing_rel_id *rel_id2)
/// {
///     return strcmp(rel_id1->target, rel_id2->target);
/// }

/// STATIC int _cond_format_hash_cmp(lxw_cond_format_hash_element *elem_1,
///                       lxw_cond_format_hash_element *elem_2)
/// {
///     return strcmp(elem_1->sqref, elem_2->sqref);
/// }

uint32_t worksheet_t::get_drawing_rel_index(const std::string& target)
{
  if(!target.empty())
  {
    const auto it = drawing_rel_ids_.find(target);
    if(it != std::end(drawing_rel_ids_))
    {
      return it->second;
    }
    else
    {
      drawing_rel_id_++;
      drawing_rel_ids_[target] = drawing_rel_id_;
      return drawing_rel_id_;
    }
  }
  else
  {
    drawing_rel_id_++;
    return drawing_rel_id_;
  }
}

uint32_t worksheet_t::find_drawing_rel_index(const std::string& target)
{
  if(target.empty())
  {
    return 0;
  }

  auto it = drawing_rel_ids_.find(target);
  if(it != std::end(drawing_rel_ids_))
  {
    return it->second;
  }
  else
  {
    return 0;
  }
}

uint32_t worksheet_t::get_vml_drawing_rel_index(const std::string& target)
{
  if(!target.empty())
  {
    const auto it = vml_drawing_rel_ids_.find(target);
    if(it != std::end(vml_drawing_rel_ids_))
    {
      return it->second;
    }
    else
    {
      vml_drawing_rel_id_++;
      vml_drawing_rel_ids_[target] = vml_drawing_rel_id_;
      return vml_drawing_rel_id_;
    }
  }
  else
  {
    vml_drawing_rel_id_++;
    return vml_drawing_rel_id_;
  }
}

uint32_t worksheet_t::find_vml_drawing_rel_index(const std::string& target)
{
  if(target.empty())
  {
    return 0;
  }

  auto it = vml_drawing_rel_ids_.find(target);
  if(it != std::end(vml_drawing_rel_ids_))
  {
    return it->second;
  }
  else
  {
    return 0;
  }
}

/// const char * lxw_basename(const char *path)
/// {
///     const char *forward_slash;
///     const char *back_slash;

///     if (!path)
///         return NULL;

///     forward_slash = strrchr(path, '/');
///     back_slash = strrchr(path, '\\');

///     if (!forward_slash && !back_slash)
///         return path;

///     if (forward_slash > back_slash)
///         return forward_slash + 1;
///     else
///         return back_slash + 1;
/// }

/// size_t _validation_list_length(const char **list)
/// {
///     uint8_t i = 0;
///     size_t length = 0;

///     if (!list || !list[0])
///         return 0;

///     while (list[i] && length < LXW_VALIDATION_MAX_STRING_LENGTH) {
/* Include commas in the length. */
///         length += 1 + lxw_utf8_strlen(list[i]);
///         i++;
///     }

/* Adjust the count for extraneous comma at end. */
///     length--;

///     return length;
/// }

/// char * _validation_list_to_csv(const char **list)
/// {
///     uint8_t i = 0;
///     char *str;

/* Create a buffer for the concatenated, and quoted, string. */
/* Allow for 4 byte UTF-8 chars and add 3 bytes for quotes and EOL. */
///     str = calloc(1, LXW_VALIDATION_MAX_STRING_LENGTH * 4 + 3);
///     if (!str)
///         return NULL;

/* Add the start quote and first element. */
///     strcat(str, "\"");
///     strcat(str, list[0]);

/* Add the other elements preceded by a comma. */
///     i = 1;
///     while (list[i]) {
///         strcat(str, ",");
///         strcat(str, list[i]);
///         i++;
///     }

/* Add the end quote. */
///     strcat(str, "\"");

///     return str;
/// }

/// STATIC double _pixels_to_width(double pixels)
/// {
///     double max_digit_width = 7.0;
///     double padding = 5.0;
///     double width;

///     if (pixels == LXW_DEF_COL_WIDTH_PIXELS)
///         width = LXW_DEF_COL_WIDTH;
///     else if (pixels <= 12.0)
///         width = pixels / (max_digit_width + padding);
///     else
///         width = (pixels - padding) / max_digit_width;

///     return width;
/// }

/// STATIC double _pixels_to_height(double pixels)
/// {
///     if (pixels == LXW_DEF_ROW_HEIGHT_PIXELS)
///         return LXW_DEF_ROW_HEIGHT;
///     else
///         return pixels * 0.75;
/// }

/// void _set_custom_filter(lxw_filter_rule_obj *rule_obj)
/// {
///     rule_obj->is_custom = LXW_TRUE;

///     if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_EQUAL_TO)
///         rule_obj->is_custom = LXW_FALSE;

///     if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
///         rule_obj->is_custom = LXW_FALSE;

///     if (rule_obj->criteria2 != LXW_FILTER_CRITERIA_NONE) {
///         if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_EQUAL_TO)
///             rule_obj->is_custom = LXW_FALSE;

///         if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
///             rule_obj->is_custom = LXW_FALSE;

///         if (rule_obj->type == LXW_FILTER_TYPE_AND)
///             rule_obj->is_custom = LXW_TRUE;
///     }

///     if (rule_obj->value1_string && strpbrk(rule_obj->value1_string, "*?"))
///         rule_obj->is_custom = LXW_TRUE;

///     if (rule_obj->value2_string && strpbrk(rule_obj->value2_string, "*?"))
///         rule_obj->is_custom = LXW_TRUE;
/// }

/// void _check_and_copy_table_style(lxw_table_obj *table_obj,
///                             lxw_table_options *user_options)
/// {
///     if (!user_options)
///         return;

/* Set the defaults. */
///     table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
///     table_obj->style_type_number = 9;

///     if (user_options->style_type > LXW_TABLE_STYLE_TYPE_DARK) {
///         LXW_WARN_FORMAT1
///             ("worksheet_add_table(): invalid style_type = %d. "
///              "Using default TableStyleMedium9", user_options->style_type);

///         table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
///         table_obj->style_type_number = 9;
///     }
///     else {
///         table_obj->style_type = user_options->style_type;
///     }

/* Each type (light, medium and dark) has a different number of styles. */
///     if (user_options->style_type == LXW_TABLE_STYLE_TYPE_LIGHT) {
///         if (user_options->style_type_number > 21) {
///             LXW_WARN_FORMAT1("worksheet_add_table(): "
///                              "invalid style_type_number = %d for style type "
///                              "LXW_TABLE_STYLE_TYPE_LIGHT. "
///                              "Using default TableStyleMedium9",
///                              user_options->style_type);

///             table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
///             table_obj->style_type_number = 9;
///         }
///         else {
///             table_obj->style_type_number = user_options->style_type_number;
///         }
///     }

///     if (user_options->style_type == LXW_TABLE_STYLE_TYPE_MEDIUM) {
///         if (user_options->style_type_number < 1
///             || user_options->style_type_number > 28) {
///             LXW_WARN_FORMAT1("worksheet_add_table(): "
///                              "invalid style_type_number = %d for style type "
///                              "LXW_TABLE_STYLE_TYPE_MEDIUM. "
///                              "Using default TableStyleMedium9",
///                              user_options->style_type_number);

///             table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
///             table_obj->style_type_number = 9;
///         }
///         else {
///             table_obj->style_type_number = user_options->style_type_number;
///         }
///     }

///     if (user_options->style_type == LXW_TABLE_STYLE_TYPE_DARK) {
///         if (user_options->style_type_number < 1
///             || user_options->style_type_number > 11) {
///             LXW_WARN_FORMAT1("worksheet_add_table(): "
///                              "invalid style_type_number = %d for style type "
///                              "LXW_TABLE_STYLE_TYPE_DARK. "
///                              "Using default TableStyleMedium9",
///                              user_options->style_type_number);

///             table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
///             table_obj->style_type_number = 9;
///         }
///         else {
///             table_obj->style_type_number = user_options->style_type_number;
///         }
///     }
/// }

/// lxw_error _set_default_table_columns(lxw_table_obj *table_obj)
/// {
///     char col_name[LXW_ATTR_32];
///     char *header;
///     uint16_t i;
///     lxw_table_column *column;
///     uint16_t num_cols = table_obj->num_cols;
///     lxw_table_column **columns = table_obj->columns;

///     for (i = 0; i < num_cols; i++) {
///         lxw_snprintf(col_name, LXW_ATTR_32, "Column%d", i + 1);

///         column = calloc(num_cols, sizeof(lxw_table_column));
///         RETURN_ON_MEM_ERROR(column, LXW_ERROR_MEMORY_MALLOC_FAILED);

///         header = lxw_strdup(col_name);
///         if (!header) {
///             free(column);
///             RETURN_ON_MEM_ERROR(header, LXW_ERROR_MEMORY_MALLOC_FAILED);
///         }
///         columns[i] = column;
///         columns[i]->header = header;
///     }

///     return LXW_NO_ERROR;
/// }

/// char * _expand_table_formula(const char *formula)
/// {
///     char *expanded;
///     const char *ptr;
///     size_t i;
///     size_t ref_count = 0;
///     size_t expanded_len;

///     ptr = formula;

///     while (*ptr) {
///         if (*ptr == '@')
///             ref_count++;

///         ptr++;
///     }

///     if (ref_count == 0) {
/* String doesn't need to be expanded. Just copy it. */
///         expanded = lxw_strdup_formula(formula);
///     }
///     else {
/* Convert "@" in the formula string to "[#This Row],".  */
///         expanded_len = strlen(formula) + (sizeof(LXW_THIS_ROW) * ref_count);
///         expanded = calloc(1, expanded_len);

///         if (!expanded)
///             return NULL;

///         i = 0;
///         ptr = formula;
/* Ignore the = in the formula. */
///         if (*ptr == '=')
///             ptr++;

/* Do the "@" expansion. */
///         while (*ptr) {
///             if (*ptr == '@') {
///                 strcat(&expanded[i], LXW_THIS_ROW);
///                 i += sizeof(LXW_THIS_ROW) - 1;
///             }
///             else {
///                 expanded[i] = *ptr;
///                 i++;
///             }

///             ptr++;
///         }
///     }

///     return expanded;
/// }

/// lxw_error _set_custom_table_columns(lxw_table_obj *table_obj,
///                           lxw_table_options *user_options)
/// {
///     char *str;
///     uint16_t i;
///     lxw_table_column *table_column;
///     lxw_table_column *user_column;
///     uint16_t num_cols = table_obj->num_cols;
///     lxw_table_column **user_columns = user_options->columns;

///     for (i = 0; i < num_cols; i++) {

///         user_column = user_columns[i];
///         table_column = table_obj->columns[i];

/* NULL indicates end of user input array. */
///         if (user_column == NULL)
///             return LXW_NO_ERROR;

///         if (user_column->header) {
///             if (lxw_utf8_strlen(user_column->header) > 255) {
///                 LXW_WARN_FORMAT("worksheet_add_table(): column parameter "
///                                 "'header' exceeds Excel length limit of 255.");
///                 return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///             }

///             str = lxw_strdup(user_column->header);
///             RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

/* Free the default column header. */
///             free((void *) table_column->header);
///             table_column->header = str;
///         }

///         if (user_column->total_string) {
///             str = lxw_strdup(user_column->total_string);
///             RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

///             table_column->total_string = str;
///         }

///         if (user_column->formula) {
///             str = _expand_table_formula(user_column->formula);
///             RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

///             table_column->formula = str;
///         }

///         table_column->format = user_column->format;
///         table_column->total_value = user_column->total_value;
///         table_column->header_format = user_column->header_format;
///         table_column->total_function = user_column->total_function;
///     }

///     return LXW_NO_ERROR;
/// }

/// void _write_column_function(lxw_worksheet *self, row_num_t row, col_num_t col,
///                        lxw_table_column *column)
/// {
///     size_t offset;
///     char formula[LXW_MAX_ATTRIBUTE_LENGTH];
///     lxw_format *format = column->format;
///     uint8_t total_function = column->total_function;
///     double value = column->total_value;
///     const char *header = column->header;

/* Write the subtotal formula number. */
///     lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH, "SUBTOTAL(%d,[",
///                  total_function);

/* Copy the header string but escape any special characters. Note, this is
 * guaranteed to fit in the 2k buffer since the header is max 255
 * characters, checked in _set_custom_table_columns(). */
///     offset = strlen(formula);
///     while (*header) {
///         switch (*header) {
///             case '\'':
///             case '#':
///             case '[':
///             case ']':
///                 formula[offset++] = '\'';
///                 formula[offset] = *header;
///                 break;
///             default:
///                 formula[offset] = *header;
///                 break;
///         }
///         offset++;
///         header++;
///     }

/* Write the end of the string. */
///     memcpy(&formula[offset], "])\0", sizeof("])\0"));

///     worksheet_write_formula_num(self, row, col, formula, format, value);
/// }

/// void _write_column_formula(lxw_worksheet *self, row_num_t first_row,
///                       row_num_t last_row, col_num_t col,
///                       lxw_table_column *column)
/// {
///     row_num_t row;
///     const char *formula = column->formula;
///     lxw_format *format = column->format;

///     for (row = first_row; row <= last_row; row++)
///         worksheet_write_formula(self, row, col, formula, format);
/// }

/// void _write_table_column_data(lxw_worksheet *self, lxw_table_obj *table_obj)
/// {
///     uint16_t i;
///     lxw_table_column *column;
///     lxw_table_column **columns = table_obj->columns;

///     col_num_t col;
///     row_num_t first_row = table_obj->first_row;
///     col_num_t first_col = table_obj->first_col;
///     row_num_t last_row = table_obj->last_row;
///     row_num_t first_data_row = first_row;
///     row_num_t last_data_row = last_row;

///     if (!table_obj->no_header_row)
///         first_data_row++;

///     if (table_obj->total_row)
///         last_data_row--;

///     for (i = 0; i < table_obj->num_cols; i++) {
///         col = first_col + i;
///         column = columns[i];

///         if (table_obj->no_header_row == LXW_FALSE)
///             worksheet_write_string(self, first_row, col, column->header,
///                                    column->header_format);

///         if (column->total_string)
///             worksheet_write_string(self, last_row, col, column->total_string,
///                                    NULL);

///         if (column->total_function)
///             _write_column_function(self, last_row, col, column);

///         if (column->formula)
///             _write_column_formula(self, first_data_row, last_data_row, col,
///                                   column);
///     }
/// }

/// lxw_error _check_table_rows(row_num_t first_row, row_num_t last_row,
///                   lxw_table_options *user_options)
/// {
///     row_num_t num_non_header_rows = last_row - first_row;

///     if (user_options && user_options->no_header_row == LXW_TRUE)
///         num_non_header_rows++;

///     if (num_non_header_rows == 0) {
///         LXW_WARN_FORMAT("worksheet_add_table(): "
///                         "table must have at least 1 non-header row.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }

///     return LXW_NO_ERROR;
/// }

/// lxw_error _check_table_name(lxw_table_options *user_options)
/// {
///     const char *name;
///     char *ptr;
///     char first[2] = { 0, 0 };

///     if (!user_options)
///         return LXW_NO_ERROR;

///     if (!user_options->name)
///         return LXW_NO_ERROR;

///     name = user_options->name;

/* Check table name length. */
///     if (lxw_utf8_strlen(name) > 255) {
///         LXW_WARN_FORMAT("worksheet_add_table(): "
///                         "Table name exceeds Excel's limit of 255.");
///         return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///     }

/* Check some short invalid names. */
///     if (strlen(name) == 1
///         && (name[0] == 'C' || name[0] == 'c' || name[0] == 'R'
///             || name[0] == 'r')) {
///         LXW_WARN_FORMAT1("worksheet_add_table(): "
///                          "invalid table name \"%s\".", name);
///         return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///     }

/* Check for invalid characters in Table name, while trying to allow
 * for utf8 strings. */
///     ptr = strpbrk(name, " !\"#$%&'()*+,-/:;<=>?@[\\]^`{|}~");
///     if (ptr) {
///         LXW_WARN_FORMAT2("worksheet_add_table(): "
///                          "invalid character '%c' in table name \"%s\".",
///                          *ptr, name);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }

/* Check for invalid initial character in Table name, while trying to allow
 * for utf8 strings. */
///     first[0] = name[0];
///     ptr = strpbrk(first, " !\"#$%&'()*+,-./0123456789:;<=>?@[\\]^`{|}~");
///     if (ptr) {
///         LXW_WARN_FORMAT2("worksheet_add_table(): "
///                          "invalid first character '%c' in table name \"%s\".",
///                          *ptr, name);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }

///     return LXW_NO_ERROR;
/// }

std::string worksheet_t::write_worksheet() const
{
  const std::vector<std::tuple<std::string, std::string>> attributes{
      {"xmlns",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main"          },
      {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
  };

  ///     if (self->excel_version == 2010) {
  ///     char xmlns_mc[] =
  ///     "http://schemas.openxmlformats.org/markup-compatibility/2006"; char
  ///     xmlns_x14ac[] =
  ///     "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
  ///         LXW_PUSH_ATTRIBUTES_STR("xmlns:mc", xmlns_mc);
  ///         LXW_PUSH_ATTRIBUTES_STR("xmlns:x14ac", xmlns_x14ac);
  ///         LXW_PUSH_ATTRIBUTES_STR("mc:Ignorable", "x14ac");
  ///     }

  return xml_start_tag("worksheet", attributes);
}

std::string worksheet_t::write_dimension() const
{
  std::string ref;

  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     char ref[LXW_MAX_CELL_RANGE_LENGTH];
  ///     row_num_t dim_rowmin = self->dim_rowmin;
  ///     row_num_t dim_rowmax = self->dim_rowmax;
  ///     col_num_t dim_colmin = self->dim_colmin;
  ///     col_num_t dim_colmax = self->dim_colmax;

  if(dim_rowmin_ == ROW_MAX && dim_colmin_ == COL_MAX)
  {
    // If the rows and cols are still the defaults then no dimensions have
    // been set and we use the default range "A1".
    ref = rowcol_to_range(0, 0, 0, 0);
  }
  else if(dim_rowmin_ == ROW_MAX && dim_colmin_ != COL_MAX)
  {
    // If the rows aren't set but the columns are then the dimensions have
    // been changed via set_column().
    ref = rowcol_to_range(0, dim_colmin_, 0, dim_colmax_);
  }
  else
  {
    ref = rowcol_to_range(dim_rowmin_, dim_colmin_, dim_rowmax_, dim_colmax_);
  }
  return xml_empty_tag("dimension", {
                                        {"ref", ref}
  });
}

/// STATIC void _worksheet_write_freeze_panes(lxw_worksheet *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     lxw_selection *selection;
///     lxw_selection *user_selection;
///     row_num_t row = self->panes.first_row;
///     col_num_t col = self->panes.first_col;
///     row_num_t top_row = self->panes.top_row;
///     col_num_t left_col = self->panes.left_col;

///     char row_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char col_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char top_left_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char active_pane[LXW_PANE_NAME_LENGTH];

/* If there is a user selection we remove it from the list and use it. */
///     if (!STAILQ_EMPTY(self->selections)) {
///         user_selection = STAILQ_FIRST(self->selections);
///         STAILQ_REMOVE_HEAD(self->selections, list_pointers);
///     }
///     else {
/* or else create a new blank selection. */
///         user_selection = calloc(1, sizeof(lxw_selection));
///         RETURN_VOID_ON_MEM_ERROR(user_selection);
///     }

///     LXW_INIT_ATTRIBUTES();

///     lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

/* Set the active pane. */
///     if (row && col) {
///         lxw_strcpy(active_pane, "bottomRight");

///         lxw_rowcol_to_cell(row_cell, row, 0);
///         lxw_rowcol_to_cell(col_cell, 0, col);

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "topRight");
///             lxw_strcpy(selection->active_cell, col_cell);
///             lxw_strcpy(selection->sqref, col_cell);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomLeft");
///             lxw_strcpy(selection->active_cell, row_cell);
///             lxw_strcpy(selection->sqref, row_cell);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomRight");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }
///     else if (col) {
///         lxw_strcpy(active_pane, "topRight");

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "topRight");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }
///     else {
///         lxw_strcpy(active_pane, "bottomLeft");

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomLeft");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }

///     if (col)
///         LXW_PUSH_ATTRIBUTES_INT("xSplit", col);

///     if (row)
///         LXW_PUSH_ATTRIBUTES_INT("ySplit", row);

///     LXW_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);
///     LXW_PUSH_ATTRIBUTES_STR("activePane", active_pane);

///     if (self->panes.type == FREEZE_PANES)
///         LXW_PUSH_ATTRIBUTES_STR("state", "frozen");
///     else if (self->panes.type == FREEZE_SPLIT_PANES)
///         LXW_PUSH_ATTRIBUTES_STR("state", "frozenSplit");

///     lxw_xml_empty_tag(self->file, "pane", &attributes);

///     free(user_selection);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC uint32_t _worksheet_calculate_x_split_width(double x_split)
/// {
///     uint32_t width;
///     uint32_t pixels;
///     uint32_t points;
///     uint32_t twips;
///     double max_digit_width = 7.0;       /* For Calabri 11. */
///     double padding = 5.0;

/* Convert to pixels. */
///     if (x_split < 1.0) {
///         pixels = (uint32_t) (x_split * (max_digit_width + padding) + 0.5);
///     }
///     else {
///         pixels = (uint32_t) (x_split * max_digit_width + 0.5) + 5;
///     }

/* Convert to points. */
///     points = (pixels * 3) / 4;

/* Convert to twips (twentieths of a point). */
///     twips = points * 20;

/* Add offset/padding. */
///     width = twips + 390;

///     return width;
/// }

/// STATIC void _worksheet_write_split_panes(lxw_worksheet *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     lxw_selection *selection;
///     lxw_selection *user_selection;
///     row_num_t row = self->panes.first_row;
///     col_num_t col = self->panes.first_col;
///     row_num_t top_row = self->panes.top_row;
///     col_num_t left_col = self->panes.left_col;
///     double x_split = self->panes.x_split;
///     double y_split = self->panes.y_split;
///     uint8_t has_selection = LXW_FALSE;

///     char row_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char col_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char top_left_cell[LXW_MAX_CELL_NAME_LENGTH];
///     char active_pane[LXW_PANE_NAME_LENGTH];

/* If there is a user selection we remove it from the list and use it. */
///     if (!STAILQ_EMPTY(self->selections)) {
///         user_selection = STAILQ_FIRST(self->selections);
///         STAILQ_REMOVE_HEAD(self->selections, list_pointers);
///         has_selection = LXW_TRUE;
///     }
///     else {
/* or else create a new blank selection. */
///         user_selection = calloc(1, sizeof(lxw_selection));
///         RETURN_VOID_ON_MEM_ERROR(user_selection);
///     }

///     LXW_INIT_ATTRIBUTES();

/* Convert the row and col to 1/20 twip units with padding. */
///     if (y_split > 0.0)
///         y_split = (uint32_t) (20 * y_split + 300);

///     if (x_split > 0.0)
///         x_split = _worksheet_calculate_x_split_width(x_split);

/* For non-explicit topLeft definitions, estimate the cell offset based on
 * the pixels dimensions. This is only a workaround and doesn't take
 * adjusted cell dimensions into account.
 */
///     if (top_row == row && left_col == col) {
///         top_row = (row_num_t) (0.5 + (y_split - 300.0) / 20.0 / 15.0);
///         left_col = (col_num_t) (0.5 + (x_split - 390.0) / 20.0 / 3.0 / 16.0);
///     }

///     lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

/* If there is no selection set the active cell to the top left cell. */
///     if (!has_selection) {
///         lxw_strcpy(user_selection->active_cell, top_left_cell);
///         lxw_strcpy(user_selection->sqref, top_left_cell);
///     }

/* Set the active pane. */
///     if (y_split > 0.0 && x_split > 0.0) {
///         lxw_strcpy(active_pane, "bottomRight");

///         lxw_rowcol_to_cell(row_cell, top_row, 0);
///         lxw_rowcol_to_cell(col_cell, 0, left_col);

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "topRight");
///             lxw_strcpy(selection->active_cell, col_cell);
///             lxw_strcpy(selection->sqref, col_cell);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomLeft");
///             lxw_strcpy(selection->active_cell, row_cell);
///             lxw_strcpy(selection->sqref, row_cell);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomRight");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }
///     else if (x_split > 0.0) {
///         lxw_strcpy(active_pane, "topRight");

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "topRight");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }
///     else {
///         lxw_strcpy(active_pane, "bottomLeft");

///         selection = calloc(1, sizeof(lxw_selection));
///         if (selection) {
///             lxw_strcpy(selection->pane, "bottomLeft");
///             lxw_strcpy(selection->active_cell, user_selection->active_cell);
///             lxw_strcpy(selection->sqref, user_selection->sqref);

///             STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///         }
///     }

///     if (x_split > 0.0)
///         LXW_PUSH_ATTRIBUTES_DBL("xSplit", x_split);

///     if (y_split > 0.0)
///         LXW_PUSH_ATTRIBUTES_DBL("ySplit", y_split);

///     LXW_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);

///     if (has_selection)
///         LXW_PUSH_ATTRIBUTES_STR("activePane", active_pane);

///     lxw_xml_empty_tag(self->file, "pane", &attributes);

///     free(user_selection);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void _worksheet_write_selection(lxw_worksheet *self, lxw_selection *selection)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (*selection->pane)
///         LXW_PUSH_ATTRIBUTES_STR("pane", selection->pane);

///     if (*selection->active_cell)
///         LXW_PUSH_ATTRIBUTES_STR("activeCell", selection->active_cell);

///     if (*selection->sqref)
///         LXW_PUSH_ATTRIBUTES_STR("sqref", selection->sqref);

///     lxw_xml_empty_tag(self->file, "selection", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void _worksheet_write_selections(lxw_worksheet *self)
/// {
///     lxw_selection *selection;

///     STAILQ_FOREACH(selection, self->selections, list_pointers) {
///         _worksheet_write_selection(self, selection);
///     }
/// }

/// STATIC void _worksheet_write_panes(lxw_worksheet *self)
/// {
///     if (self->panes.type == NO_PANES)
///         return;

///     else if (self->panes.type == FREEZE_PANES)
///         _worksheet_write_freeze_panes(self);

///     else if (self->panes.type == FREEZE_SPLIT_PANES)
///         _worksheet_write_freeze_panes(self);
///
///     else if (self->panes.type == SPLIT_PANES)
///         _worksheet_write_split_panes(self);
/// }

std::string worksheet_t::write_sheet_view() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  /* Hide screen gridlines if required */
  ///     if (!self->screen_gridlines)
  ///         LXW_PUSH_ATTRIBUTES_STR("showGridLines", "0");

  /* Hide zeroes in cells. */
  ///     if (!self->show_zeros)
  ///         LXW_PUSH_ATTRIBUTES_STR("showZeros", "0");

  /* Display worksheet right to left for Hebrew, Arabic and others. */
  ///     if (self->right_to_left)
  ///         LXW_PUSH_ATTRIBUTES_STR("rightToLeft", "1");

  // Show that the sheet tab is selected.
  if(selected_)
  {
    attributes.emplace_back("tabSelected", "1");
  }

  /* Turn outlines off. Also required in the outlinePr element. */
  ///     if (!self->outline_on)
  ///         LXW_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

  /* Set the page view/layout mode if required. */
  ///     if (self->page_view)
  ///         LXW_PUSH_ATTRIBUTES_STR("view", "pageLayout");

  /* Set the top left cell if required. */
  ///     if (self->top_left_cell[0])
  ///         LXW_PUSH_ATTRIBUTES_STR("topLeftCell", self->top_left_cell);

  /* Set the zoom level. */
  ///     if (self->zoom != 100 && !self->page_view) {
  ///         LXW_PUSH_ATTRIBUTES_INT("zoomScale", self->zoom);

  ///         if (self->zoom_scale_normal)
  ///             LXW_PUSH_ATTRIBUTES_INT("zoomScaleNormal", self->zoom);
  ///     }

  attributes.emplace_back("workbookViewId", "0");

  ///     if (self->panes.type != NO_PANES || !STAILQ_EMPTY(self->selections)) {
  ///         lxw_xml_start_tag(self->file, "sheetView", &attributes);
  ///         _worksheet_write_panes(self);
  ///         _worksheet_write_selections(self);
  ///         lxw_xml_end_tag(self->file, "sheetView");
  ///     }
  ///     else {
  return xml_empty_tag("sheetView", attributes);
  ///     }
}

std::string worksheet_t::write_sheet_views() const
{
  std::string xml_data = xml_start_tag("sheetViews");
  xml_data += write_sheet_view();
  xml_data += xml_end_tag("sheetViews");

  return xml_data;
}

std::string worksheet_t::write_sheet_format_pr() const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"defaultRowHeight", std::format("{}", default_row_height_)}
  };

  if(default_row_height_ != DEF_ROW_HEIGHT)
  {
    attributes.emplace_back("customHeight", "1");
  }

  ///     if (self->default_row_zeroed)
  ///         LXW_PUSH_ATTRIBUTES_STR("zeroHeight", "1");

  ///     if (self->outline_row_level)
  ///         LXW_PUSH_ATTRIBUTES_INT("outlineLevelRow",
  ///         self->outline_row_level);

  ///     if (self->outline_col_level)
  ///         LXW_PUSH_ATTRIBUTES_INT("outlineLevelCol",
  ///         self->outline_col_level);

  ///     if (self->excel_version == 2010)
  ///         LXW_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

  return xml_empty_tag("sheetFormatPr", attributes);
}

std::string worksheet_t::write_sheet_data() const
{

  if(table_.rbh_root_.empty())
  {
    return xml_empty_tag("sheetData");
  }
  else
  {
    std::string xml_data = xml_start_tag("sheetData");
    xml_data += write_rows();
    xml_data += xml_end_tag("sheetData");
    return xml_data;
  }
}

/// STATIC void _worksheet_write_optimized_sheet_data(lxw_worksheet *self)
/// {
///     size_t read_size = 1;
///     char buffer[LXW_BUFFER_SIZE];

///     if (self->dim_rowmin == LXW_ROW_MAX) {
/* If the dimensions aren't defined then there is no data to write. */
///         lxw_xml_empty_tag(self->file, "sheetData", NULL);
///     }
///     else {
///         lxw_xml_start_tag(self->file, "sheetData", NULL);

/* Flush the temp file. */
///         fflush(self->optimize_tmpfile);

///         if (self->optimize_buffer) {
/* Ignore return value. There is no easy way to raise error. */
///             (void) fwrite(self->optimize_buffer, self->optimize_buffer_size,
///                           1, self->file);
///         }
///         else {
/* Rewind the temp file. */
///             rewind(self->optimize_tmpfile);
///             while (read_size) {
///                 read_size =
///                     fread(buffer, 1, LXW_BUFFER_SIZE, self->optimize_tmpfile);
/* Ignore return value. There is no easy way to raise error. */
///                 (void) fwrite(buffer, 1, read_size, self->file);
///             }
///         }

///         fclose(self->optimize_tmpfile);
///         free(self->optimize_buffer);

///         lxw_xml_end_tag(self->file, "sheetData");
///     }
/// }

std::string worksheet_t::write_page_margins() const
{
  return xml_empty_tag("pageMargins", {
                                          {"left",   std::format("{}", margin_left_)  },
                                          {"right",  std::format("{}", margin_right_) },
                                          {"top",    std::format("{}", margin_top_)   },
                                          {"bottom", std::format("{}", margin_bottom_)},
                                          {"header", std::format("{}", margin_header_)},
                                          {"footer", std::format("{}", margin_footer_)},
  });
}

/*
 * Write the <pageSetup> element.
 * The following is an example taken from Excel.
 * <pageSetup
 *     paperSize="9"
 *     scale="110"
 *     fitToWidth="2"
 *     fitToHeight="2"
 *     pageOrder="overThenDown"
 *     orientation="portrait"
 *     blackAndWhite="1"
 *     draft="1"
 *     horizontalDpi="200"
 *     verticalDpi="200"
 *     r:id="rId1"
 * />
 */
std::string worksheet_t::write_page_setup() const
{
  std::string xml_data;
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;

  ///     if (!self->page_setup_changed)
  ///         return;

  /* Set paper size. */
  ///     if (self->paper_size)
  ///         LXW_PUSH_ATTRIBUTES_INT("paperSize", self->paper_size);

  /* Set the print_scale. */
  ///     if (self->print_scale != 100)
  ///         LXW_PUSH_ATTRIBUTES_INT("scale", self->print_scale);

  /* Set the "Fit to page" properties. */
  ///     if (self->fit_page && self->fit_width != 1)
  ///         LXW_PUSH_ATTRIBUTES_INT("fitToWidth", self->fit_width);

  ///     if (self->fit_page && self->fit_height != 1)
  ///         LXW_PUSH_ATTRIBUTES_INT("fitToHeight", self->fit_height);

  /* Set the page print direction. */
  ///     if (self->page_order)
  ///         LXW_PUSH_ATTRIBUTES_STR("pageOrder", "overThenDown");

  /* Set start page. */
  ///     if (self->page_start > 1)
  ///         LXW_PUSH_ATTRIBUTES_INT("firstPageNumber", self->page_start);

  /* Set page orientation. */
  ///     if (self->orientation)
  ///         LXW_PUSH_ATTRIBUTES_STR("orientation", "portrait");
  ///     else
  ///         LXW_PUSH_ATTRIBUTES_STR("orientation", "landscape");

  ///     if (self->black_white)
  ///         LXW_PUSH_ATTRIBUTES_STR("blackAndWhite", "1");

  /* Set start page active flag. */
  ///     if (self->page_start)
  ///         LXW_PUSH_ATTRIBUTES_INT("useFirstPageNumber", 1);

  /* Set the DPI. Mainly only for testing. */
  ///     if (self->horizontal_dpi)
  ///         LXW_PUSH_ATTRIBUTES_INT("horizontalDpi", self->horizontal_dpi);

  ///     if (self->vertical_dpi)
  ///         LXW_PUSH_ATTRIBUTES_INT("verticalDpi", self->vertical_dpi);

  ///     lxw_xml_empty_tag(self->file, "pageSetup", &attributes);

  return xml_data;
}

std::string worksheet_t::write_print_options() const
{
  std::string xml_data;

  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     if (!self->print_options_changed)
  ///         return;

  ///     /* Set horizontal centering. */
  ///     if (self->hcenter) {
  ///         LXW_PUSH_ATTRIBUTES_STR("horizontalCentered", "1");
  ///     }

  ///     /* Set vertical centering. */
  ///     if (self->vcenter) {
  ///         LXW_PUSH_ATTRIBUTES_STR("verticalCentered", "1");
  ///     }

  /* Enable row and column headers. */
  ///     if (self->print_headers) {
  ///         LXW_PUSH_ATTRIBUTES_STR("headings", "1");
  ///     }

  /* Set printed gridlines. */
  ///     if (self->print_gridlines) {
  ///         LXW_PUSH_ATTRIBUTES_STR("gridLines", "1");
  ///     }

  ///     lxw_xml_empty_tag(self->file, "printOptions", &attributes);

  return xml_data;
}

std::string worksheet_t::write_row(const row_t& row) const
{
  return write_row(row, "");
}

std::string worksheet_t::write_row(const row_t& row, const std::string& spans) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  ///     int32_t xf_index = 0;
  const double height = (row.height_changed_ ? row.height_ : default_row_height_);

  ///     if (row->format) {
  ///         xf_index = lxw_format_get_xf_index(row->format);
  ///     }

  attributes.emplace_back("r", std::to_string(row.row_num_ + 1));

  if(!spans.empty())
  {
    attributes.emplace_back("spans", spans);
  }

  ///     if (xf_index)
  ///         LXW_PUSH_ATTRIBUTES_INT("s", xf_index);

  ///     if (row->format)
  ///         LXW_PUSH_ATTRIBUTES_STR("customFormat", "1");

  if(height != DEF_ROW_HEIGHT)
  {
    attributes.emplace_back("ht", std::format("{}", height));
  }

  ///     if (row->hidden)
  ///         LXW_PUSH_ATTRIBUTES_STR("hidden", "1");

  if(height != DEF_ROW_HEIGHT)
  {
    attributes.emplace_back("customHeight", "1");
  }

  ///     if (row->level)
  ///         LXW_PUSH_ATTRIBUTES_INT("outlineLevel", row->level);

  ///     if (row->collapsed)
  ///         LXW_PUSH_ATTRIBUTES_STR("collapsed", "1");

  ///     if (self->excel_version == 2010)
  ///         LXW_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

  if(!row.data_changed_)
  {
    return xml_empty_tag("row", attributes);
  }
  else
  {
    return xml_start_tag("row", attributes);
  }
}

int32_t worksheet_t::size_col(col_num_t col_num, object_position_t anchor)
{
  col_options_t* col_opt = nullptr;
  uint32_t pixels        = 0;

  // Search for the col number in the array of col_options. Each col_option
  // entry contains the start and end column for a range.
  for(col_num_t col_index = 0; col_index < col_options_max_; col_index++)
  {
    if(col_index < col_options_.size())
    {
      col_opt = &col_options_[col_index];
      if(col_num >= col_opt->firstcol_ && col_num <= col_opt->lastcol_)
      {
        break;
      }
      else
      {
        col_opt = nullptr;
      }
    }
  }

  if(col_opt)
  {
    double width                 = col_opt->width_;
    const double max_digit_width = 7.0; /* For Calabri 11. */
    const double padding         = 5.0;

    // Convert to pixels.
    if(col_opt->hidden_ && anchor != object_position_t::MOVE_AND_SIZE_AFTER)
    {
      pixels = 0;
    }
    else if(width < 1.0)
    {
      pixels = (uint32_t)(width * (max_digit_width + padding) + 0.5);
    }
    else
    {
      pixels = (uint32_t)(width * max_digit_width + 0.5) + 5;
    }
  }
  else
  {
    pixels = default_col_pixels_;
  }

  return pixels;
}

/*
 * Convert the height of a cell from user's units to pixels. If the height
 * hasn't been set by the user we use the default value. If the row is hidden
 * it has a value of zero.
 */
int32_t worksheet_t::size_row(row_num_t row_num, object_position_t anchor)
{
  uint32_t pixels = 0;

  const row_t* row = find_row(row_num);
  /* Note, the 0.75 below is due to the difference between 72/96 DPI. */
  if(row)
  {
    if(row->hidden_ && anchor != object_position_t::MOVE_AND_SIZE_AFTER)
    {
      pixels = 0;
    }
    else
    {
      pixels = static_cast<uint32_t>(row->height_ / 0.75);
    }
  }
  else
  {
    pixels = static_cast<uint32_t>(default_row_height_ / 0.75);
  }

  return pixels;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in pixels.
 *         +------------+------------+
 *         |     A      |      B     |
 *   +-----+------------+------------+
 *   |     |(x1,y1)     |            |
 *   |  1  |(A1)._______|______      |
 *   |     |    |              |     |
 *   |     |    |              |     |
 *   +-----+----|    BITMAP    |-----+
 *   |     |    |              |     |
 *   |  2  |    |______________.     |
 *   |     |            |        (B2)|
 *   |     |            |     (x2,y2)|
 *   +---- +------------+------------+
 *
 * Example of an object that covers some of the area from cell A1 to cell B2.
 * Based on the width and height of the object we need to calculate 8 vars:
 *
 *     col_start, row_start, col_end, row_end, x1, y1, x2, y2.
 *
 * We also calculate the absolute x and y position of the top left vertex of
 * the object. This is required for images:
 *
 *    x_abs, y_abs
 *
 * The width and height of the cells that the object occupies can be variable
 * and have to be taken into account.
 *
 * The values of col_start and row_start are passed in from the calling
 * function. The values of col_end and row_end are calculated by subtracting
 * the width and height of the object from the width and height of the
 * underlying cells.
 */
void worksheet_t::position_object_pixels(const object_properties_t& object_props, drawing_object_t& drawing_object)
{
  col_num_t col_start = object_props.col_;      // Column containing upper left corner.
  int32_t x1          = object_props.x_offset_; // Distance to left side of object.
  row_num_t row_start = object_props.row_;      // Row containing top left corner.
  int32_t y1          = object_props.y_offset_; // Distance to top of object.
  col_num_t col_end;                            // Column containing lower right corner.
  double x2;                                    // Distance to right side of object.
  row_num_t row_end;                            // Row containing bottom right corner.
  double y2;                                    // Distance to bottom of object.
  double width   = object_props.width_;         // Width of object frame.
  double height  = object_props.height_;        // Height of object frame.
  uint32_t x_abs = 0;                           // Abs. distance to left side of object.
  uint32_t y_abs = 0;                           // Abs. distance to top  side of object.
  uint32_t i;
  object_position_t anchor        = static_cast<object_position_t>(drawing_object.anchor_);
  object_position_t ignore_anchor = object_position_t::DEFAULT;

  // Adjust start column for negative offsets.
  while(x1 < 0 && col_start > 0)
  {
    x1 += size_col(col_start - 1, ignore_anchor);
    col_start--;
  }

  // Adjust start row for negative offsets.
  while(y1 < 0 && row_start > 0)
  {
    y1 += size_row(row_start - 1, ignore_anchor);
    row_start--;
  }

  // Ensure that the image isn't shifted off the page at top left.
  if(x1 < 0)
  {
    x1 = 0;
  }

  if(y1 < 0)
  {
    y1 = 0;
  }

  // Calculate the absolute x offset of the top-left vertex.
  if(col_size_changed_)
  {
    for(i = 0; i < col_start; i++)
    {
      x_abs += size_col(i, ignore_anchor);
    }
  }
  else
  {
    // Optimization for when the column widths haven't changed.
    x_abs += default_col_pixels_ * col_start;
  }
  x_abs += x1;

  // Calculate the absolute y offset of the top-left vertex.
  // Store the column change to allow optimizations.
  if(row_size_changed_)
  {
    for(i = 0; i < row_start; i++)
    {
      y_abs += size_row(i, ignore_anchor);
    }
  }
  else
  {
    // Optimization for when the row heights haven"t changed.
    y_abs += default_row_pixels_ * row_start;
  }
  y_abs += y1;

  // Adjust start col for offsets that are greater than the col width.
  while(x1 >= size_col(col_start, anchor))
  {
    x1 -= size_col(col_start, ignore_anchor);
    col_start++;
  }

  // Adjust start row for offsets that are greater than the row height.
  while(y1 >= size_row(row_start, anchor))
  {
    y1 -= size_row(row_start, ignore_anchor);
    row_start++;
  }

  // Initialize end cell to the same as the start cell.
  col_end = col_start;
  row_end = row_start;

  // Only offset the image in the cell if the row/col is hidden.
  if(size_col(col_start, anchor) > 0)
  {
    width = width + x1;
  }
  if(size_row(row_start, anchor) > 0)
  {
    height = height + y1;
  }

  // Subtract the underlying cell widths to find the end cell.
  while(width >= size_col(col_end, anchor) && col_end < COL_MAX)
  {
    width -= size_col(col_end, anchor);
    col_end++;
  }

  // Subtract the underlying cell heights to find the end cell.
  while(height >= size_row(row_end, anchor) && row_end < ROW_MAX)
  {
    height -= size_row(row_end, anchor);
    row_end++;
  }

  // The end vertices are whatever is left from the width and height.
  x2 = width;
  y2 = height;

  // Add the dimensions to the drawing object.
  drawing_object.from_.col_        = col_start;
  drawing_object.from_.row_        = row_start;
  drawing_object.from_.col_offset_ = x1;
  drawing_object.from_.row_offset_ = y1;
  drawing_object.to_.col_          = col_end;
  drawing_object.to_.row_          = row_end;
  drawing_object.to_.col_offset_   = x2;
  drawing_object.to_.row_offset_   = y2;
  drawing_object.col_absolute_     = x_abs;
  drawing_object.row_absolute_     = y_abs;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in EMUs. The vertices are expressed as English
 * Metric Units (EMUs). There are 12,700 EMUs per point.
 * Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
 */
void worksheet_t::position_object_emus(const object_properties_t& image, drawing_object_t& drawing_object)
{
  position_object_pixels(image, drawing_object);

  // Convert the pixel values to EMUs. See above.
  drawing_object.from_.col_offset_ *= 9525;
  drawing_object.from_.row_offset_ *= 9525;
  drawing_object.to_.col_offset_ *= 9525;
  drawing_object.to_.row_offset_ *= 9525;
  drawing_object.to_.col_offset_ += 0.5;
  drawing_object.to_.row_offset_ += 0.5;
  drawing_object.col_absolute_ *= 9525;
  drawing_object.row_absolute_ *= 9525;
}

/*
 * This function handles the additional optional parameters to
 * worksheet_insert_button() as well as calculating the button object
 * position and vertices.
 */
/// lxw_error
/// _get_button_params(lxw_vml_obj *button, uint16_t button_number,
///                    lxw_button_options *options)
/// {
///     int32_t x_offset = 0;
///     int32_t y_offset = 0;
///     uint32_t height = LXW_DEF_ROW_HEIGHT_PIXELS;
///     uint32_t width = LXW_DEF_COL_WIDTH_PIXELS;
///     double x_scale = 1.0;
///     double y_scale = 1.0;
///     row_num_t row = button->row;
///     col_num_t col = button->col;
///     char buffer[LXW_ATTR_32];
///     uint8_t has_caption = LXW_FALSE;
///     uint8_t has_macro = LXW_FALSE;
///     size_t len;

/* Set any user defined options. */
///     if (options) {
///         if (options->width > 0.0)
///             width = options->width;

///         if (options->height > 0.0)
///             height = options->height;

///         if (options->x_scale > 0.0)
///             x_scale = options->x_scale;

///         if (options->y_scale > 0.0)
///             y_scale = options->y_scale;

///         if (options->x_offset != 0)
///             x_offset = options->x_offset;

///         if (options->y_offset != 0)
///             y_offset = options->y_offset;

///         if (options->caption) {
///             button->name = lxw_strdup(options->caption);
///             RETURN_ON_MEM_ERROR(button->name, LXW_ERROR_MEMORY_MALLOC_FAILED);
///             has_caption = LXW_TRUE;
///         }

///         if (options->macro) {
///             len = sizeof("[0]!") + strlen(options->macro);
///             button->macro = calloc(1, len);
///             RETURN_ON_MEM_ERROR(button->macro,
///                                 LXW_ERROR_MEMORY_MALLOC_FAILED);

///             if (button->macro)
///                 lxw_snprintf(button->macro, len, "[0]!%s", options->macro);

///             has_macro = LXW_TRUE;
///         }

///         if (options->description) {
///             button->text = lxw_strdup(options->description);
///             RETURN_ON_MEM_ERROR(button->text, LXW_ERROR_MEMORY_MALLOC_FAILED);
///         }
///     }

///     if (!has_caption) {
///         lxw_snprintf(buffer, LXW_ATTR_32, "Button %d", button_number);
///         button->name = lxw_strdup(buffer);
///         RETURN_ON_MEM_ERROR(button->name, LXW_ERROR_MEMORY_MALLOC_FAILED);
///     }

///     if (!has_macro) {
///         lxw_snprintf(buffer, LXW_ATTR_32, "[0]!Button%d_Click",
///                      button_number);
///         button->macro = lxw_strdup(buffer);
///         RETURN_ON_MEM_ERROR(button->macro, LXW_ERROR_MEMORY_MALLOC_FAILED);
///     }

/* Scale the width/height to the default/user scale and round to the
 * nearest pixel. */
///     width = (uint32_t) (0.5 + x_scale * width);
///     height = (uint32_t) (0.5 + y_scale * height);

///     button->width = width;
///     button->height = height;
///     button->start_col = col;
///     button->start_row = row;
///     button->x_offset = x_offset;
///     button->y_offset = y_offset;

///     return LXW_NO_ERROR;
/// }

void worksheet_t::position_vml_object(vml_obj_t& vml_obj)
{
  object_properties_t object_props;
  drawing_object_t drawing_object;

  object_props.col_      = vml_obj.start_col_;
  object_props.row_      = vml_obj.start_row_;
  object_props.x_offset_ = vml_obj.x_offset_;
  object_props.y_offset_ = vml_obj.y_offset_;
  object_props.width_    = vml_obj.width_;
  object_props.height_   = vml_obj.height_;

  drawing_object.anchor_ = static_cast<uint8_t>(object_position_t::DONT_MOVE_DONT_SIZE);

  position_object_pixels(object_props, drawing_object);

  vml_obj.from_.col_        = drawing_object.from_.col_;
  vml_obj.from_.row_        = drawing_object.from_.row_;
  vml_obj.from_.col_offset_ = drawing_object.from_.col_offset_;
  vml_obj.from_.row_offset_ = drawing_object.from_.row_offset_;
  vml_obj.to_.col_          = drawing_object.to_.col_;
  vml_obj.to_.row_          = drawing_object.to_.row_;
  vml_obj.to_.col_offset_   = drawing_object.to_.col_offset_;
  vml_obj.to_.row_offset_   = drawing_object.to_.row_offset_;
  vml_obj.col_absolute_     = drawing_object.col_absolute_;
  vml_obj.row_absolute_     = drawing_object.row_absolute_;
}

void worksheet_t::prepare_image(uint32_t image_ref_id, uint32_t drawing_id, object_properties_t& object_props)
{
  if(!drawing_)
  {
    drawing_ = drawing_t{};
    external_drawing_links_.emplace_back("/drawing", std::format("../drawings/drawing{}.xml", drawing_id), "");
  }

  drawing_object_t drawing_object;
  drawing_object.anchor_ = static_cast<uint8_t>(object_position_t::MOVE_DONT_SIZE);
  if(object_props.object_position_ != object_position_t::DEFAULT)
  {
    drawing_object.anchor_ = static_cast<uint8_t>(object_props.object_position_);
  }

  drawing_object.type_          = drawing_types_t::IMAGE;
  drawing_object.description_   = object_props.description_;
  drawing_object.tip_           = object_props.tip_;
  drawing_object.rel_index_     = 0;
  drawing_object.url_rel_index_ = 0;
  drawing_object.decorative_    = object_props.decorative_;

  // Scale to user scale.
  double width  = object_props.width_ * object_props.x_scale_;
  double height = object_props.height_ * object_props.y_scale_;

  // Scale by non 96dpi resolutions.
  width *= 96.0 / object_props.x_dpi_;
  height *= 96.0 / object_props.y_dpi_;

  object_props.width_  = width;
  object_props.height_ = height;

  position_object_emus(object_props, drawing_object);

  // Convert from pixels to emus.
  drawing_object.width_  = static_cast<uint32_t>(0.5 + width * 9525);
  drawing_object.height_ = static_cast<uint32_t>(0.5 + height * 9525);

  if(!object_props.url_.empty())
  {
    std::string url = object_props.url_;

    // Check the link type. Default to external hyperlinks.
    // TODO Remove this useless variable
    cell_types_t link_type = cell_types_t::HYPERLINK_URL;
    if(url.find("internal:") != std::string::npos)
    {
      link_type = cell_types_t::HYPERLINK_INTERNAL;
    }
    else if(url.find("external:") != std::string::npos)
    {
      link_type = cell_types_t::HYPERLINK_EXTERNAL;
    }

    // Set the relationship object for each type of link.
    if(link_type == cell_types_t::HYPERLINK_INTERNAL)
    {
      const std::string target = "#"s + url.substr(sizeof("internal"));
      drawing_links_.emplace_back("/hyperlink", target, "");
    }
    else if(link_type == cell_types_t::HYPERLINK_EXTERNAL)
    {
      std::string url_copy = url.substr(sizeof("external"));

      // Look for Windows style "C:/" link or Windows share "\\" link.
      size_t found_string = url_copy.find(':');
      if(found_string == std::string::npos)
      {
        found_string = url_copy.find("\\\\");
      }

      if(found_string != std::string::npos)
      {
        // Add the file:/// URI to the url if non-local.
        drawing_links_.emplace_back("/hyperlink", "file:///"s + escape_url_characters(url_copy, true), "External");
      }
      else
      {
        // Copy the relative url without "external:".
        std::string url_copy = escape_url_characters(url.substr(sizeof("external")), true);

        // Switch backslash to forward slash.
        for(auto& c: url_copy)
        {
          if(c == '/')
          {
            c = '\\';
          }
        }
        drawing_links_.emplace_back("/hyperlink", url_copy, "External");
      }
    }
    else
    {
      drawing_links_.emplace_back("/hyperlink", escape_url_characters(object_props.url_, false), "External");
    }

    drawing_object.url_rel_index_ = get_drawing_rel_index(url);
  }

  if(find_drawing_rel_index(object_props.md5_) == 0)
  {
    drawing_links_.emplace_back("/image", std::format("../media/image{}.{}", image_ref_id, object_props.extension_),
                                "");
  }

  drawing_object.rel_index_ = get_drawing_rel_index(object_props.md5_);
  drawing_->add_drawing_object(drawing_object);
}

void worksheet_t::prepare_header_image(uint32_t image_ref_id, object_properties_t& object_props)
{
  ///     lxw_rel_tuple *relationship = NULL;
  ///     char filename[LXW_FILENAME_LENGTH];
  ///     lxw_vml_obj *header_image_vml;
  ///     char *extension;

  image_props_.push_back(object_props);

  if(find_vml_drawing_rel_index(object_props.md5_) == 0)
  {
    vml_drawing_links_.emplace_back("/image", std::format("../media/image{}.{}", image_ref_id, object_props.extension_),
                                    "");
  }

  vml_obj_t header_image_vml;
  header_image_vml.width_          = static_cast<uint32_t>(object_props.width_);
  header_image_vml.height_         = static_cast<uint32_t>(object_props.height_);
  header_image_vml.x_dpi_          = object_props.x_dpi_;
  header_image_vml.y_dpi_          = object_props.y_dpi_;
  header_image_vml.rel_index_      = 1;
  header_image_vml.image_position_ = object_props.image_position_;
  header_image_vml.name_           = object_props.description_;

  // Strip the extension from the filename.
  const size_t pos = header_image_vml.name_.find_last_of('.');
  if(pos != std::string::npos)
  {
    header_image_vml.name_ = header_image_vml.name_.substr(0, pos);
  }

  header_image_vml.rel_index_ = get_vml_drawing_rel_index(object_props.md5_);
  header_image_objs_.push_back(header_image_vml);
}

/// void
/// lxw_worksheet_prepare_background(lxw_worksheet *self,
///                                  uint32_t image_ref_id,
///                                  lxw_object_properties *object_props)
/// {
///     lxw_rel_tuple *relationship = NULL;
///     char filename[LXW_FILENAME_LENGTH];

///     STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);

///     relationship = calloc(1, sizeof(lxw_rel_tuple));
///     RETURN_VOID_ON_MEM_ERROR(relationship);

///     relationship->type = lxw_strdup("/image");
///     GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

///     lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
///                  object_props->extension);

///     relationship->target = lxw_strdup(filename);
///     GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

///     self->external_background_link = relationship;

///     return;

/// mem_error:
///     if (relationship) {
///         free(relationship->type);
///         free(relationship->target);
///         free(relationship->target_mode);
///         free(relationship);
///     }
/// }

void worksheet_t::prepare_chart(uint32_t chart_ref_id, uint32_t drawing_id, object_properties_t& object_props,
                                bool is_chartsheet)
{
  if(!drawing_)
  {
    drawing_ = drawing_t{};
    if(is_chartsheet)
    {
      drawing_->embedded_    = false;
      drawing_->orientation_ = orientation_;
    }
    else
    {
      drawing_->embedded_ = true;
    }
    external_drawing_links_.emplace_back("/drawing", std::format("../drawings/drawing{}.xml", drawing_id), "");
  }

  drawing_object_t drawing_object;
  drawing_object.anchor_ = static_cast<uint8_t>(object_position_t::MOVE_AND_SIZE);
  if(object_props.object_position_ != object_position_t::DEFAULT)
  {
    drawing_object.anchor_ = static_cast<uint8_t>(object_props.object_position_);
  }

  drawing_object.type_          = drawing_types_t::CHART;
  drawing_object.description_   = object_props.description_;
  drawing_object.tip_           = "";
  drawing_object.rel_index_     = get_drawing_rel_index("");
  drawing_object.url_rel_index_ = 0;
  drawing_object.decorative_    = object_props.decorative_;

  // Scale to user scale.
  double width  = object_props.width_ * object_props.x_scale_;
  double height = object_props.height_ * object_props.y_scale_;

  // Convert to the nearest pixel.
  object_props.width_  = width;
  object_props.height_ = height;

  position_object_emus(object_props, drawing_object);

  // Convert from pixels to emus.
  drawing_object.width_  = static_cast<uint32_t>(0.5 + width * 9525);
  drawing_object.height_ = static_cast<uint32_t>(0.5 + height * 9525);

  drawing_->add_drawing_object(drawing_object);
  drawing_links_.emplace_back("/chart", std::format("../charts/chart{}.xml", chart_ref_id), "");
}

uint32_t worksheet_t::prepare_vml_objects(uint32_t vml_data_id, uint32_t vml_shape_id, uint32_t vml_drawing_id,
                                          uint32_t comment_id)
{
  uint32_t comment_count = 0;

  for(auto& [index, row]: comments_.rbh_root_)
  {
    for(auto& [index, cell]: row.cells_)
    {
      // Calculate the worksheet position of the comment.
      position_vml_object(cell.comment_.value());

      // Store comment in a simple list for use by packager.
      comment_objs_.push_back(cell.comment_.value());
      comment_count++;
    }
  }

  // Set up the VML relationship for comments/buttons/header images.
  external_vml_comment_link_ =
      std::make_tuple("/vmlDrawing"s, std::format("../drawings/vmlDrawing{}.vml", vml_drawing_id), ""s);

  if(has_comments_)
  {
    // Only need this relationship object for comment VMLs.
    external_comment_link_ = std::make_tuple("/comments"s, std::format("../comments{}.xml", comment_id), ""s);
  }

  // Create the CSV list in the allocated space.
  for(size_t i = 0; i <= comment_count / 1024; i++)
  {
    vml_data_id_str_ += std::format("{},", vml_data_id + i);
  }

  // Remove last comma
  if(!vml_data_id_str_.empty() && vml_data_id_str_.back() == ',')
  {
    vml_data_id_str_.pop_back();
  }

  vml_shape_id_ = vml_shape_id;

  return comment_count;
}

void worksheet_t::prepare_header_vml_objects(uint32_t vml_header_id, uint32_t vml_drawing_id)
{
  vml_header_id_ = vml_header_id;

  // Set up the VML relationship for header images.
  external_vml_header_link_ =
      std::make_tuple("/vmlDrawing"s, std::format("../drawings/vmlDrawing{}.vml", vml_drawing_id), ""s);

  vml_header_id_str_ = std::to_string(vml_header_id);
}

/// void lxw_worksheet_prepare_tables(lxw_worksheet *self, uint32_t table_id)
/// {
///     lxw_table_obj *table_obj;
///     lxw_rel_tuple *relationship;
///     char name[LXW_ATTR_32];
///     char filename[LXW_FILENAME_LENGTH];

///     STAILQ_FOREACH(table_obj, self->table_objs, list_pointers) {

///         relationship = calloc(1, sizeof(lxw_rel_tuple));
///         GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

///         relationship->type = lxw_strdup("/table");
///         GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

///         lxw_snprintf(filename, LXW_FILENAME_LENGTH,
///                      "../tables/table%d.xml", table_id);

///         relationship->target = lxw_strdup(filename);
///         GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

///         STAILQ_INSERT_TAIL(self->external_table_links, relationship,
///                            list_pointers);

///         if (!table_obj->name) {
///             lxw_snprintf(name, LXW_ATTR_32, "Table%d", table_id);
///             table_obj->name = lxw_strdup(name);
///             GOTO_LABEL_ON_MEM_ERROR(table_obj->name, mem_error);
///         }
///         table_obj->id = table_id;
///         table_id++;
///     }

///     return;

/// mem_error:
///     if (relationship) {
///         free(relationship->type);
///         free(relationship->target);
///         free(relationship->target_mode);
///         free(relationship);
///     }

///     return;
/// }

void process_png(object_properties_t& image_props, const std::vector<unsigned char>& data)
{
  uint32_t width  = 0;
  uint32_t height = 0;
  double x_dpi    = 96;
  double y_dpi    = 96;

  // Start after header;
  std::vector<unsigned char>::const_iterator it = std::begin(data);
  it += (4 + 4);

  while(it + 3 * 4 < std::end(data))
  {
    // Read the PNG length and type fields for the sub-section.
    const uint32_t length = (it[0] << 24) + (it[1] << 16) + (it[2] << 8) + it[3];
    it += 4;

    std::string type{it, it + 4};
    it += 4;

    if(it + length < std::end(data))
    {
      if(type == "IHDR")
      {
        width = (it[0] << 24) + (it[1] << 16) + (it[2] << 8) + it[3];
        it += 4;

        height = (it[0] << 24) + (it[1] << 16) + (it[2] << 8) + it[3];
        it += 4;

        it += length - 2 * 4;
      }
      else if(type == "pHYs")
      {
        const uint32_t x_ppu = (it[0] << 24) + (it[1] << 16) + (it[2] << 8) + it[3];
        it += 4;

        const uint32_t y_ppu = (it[0] << 24) + (it[1] << 16) + (it[2] << 8) + it[3];
        it += 4;

        const uint8_t units = it[0];
        it++;

        if(units == 1)
        {
          x_dpi = x_ppu * 0.0254;
          y_dpi = y_ppu * 0.0254;
        }

        it += length - (2 * 4 + 1);
      }
      else if(type == "IEND")
      {
        break;
      }
      else
      {
        it += length;
      }
    }
    else
    {
      // No enough byte for subsection, ==> stop
      break;
    }

    // Ignore CRC
    it += 4;
  }

  // Set the image metadata.
  image_props.image_type_ = image_types_t::PNG;
  image_props.width_      = width;
  image_props.height_     = height;
  image_props.x_dpi_      = x_dpi;
  image_props.y_dpi_      = y_dpi;
  image_props.extension_  = "png";
}

/// STATIC lxw_error
/// _process_jpeg(lxw_object_properties *image_props)
/// {
///     uint16_t length;
///     uint16_t marker;
///     uint32_t offset;
///     uint16_t width = 0;
///     uint16_t height = 0;
///     double x_dpi = 96;
///     double y_dpi = 96;
///     int fseek_err;

///     FILE *stream = image_props->stream;

/* Read back 2 bytes to the end of the initial 0xFFD8 marker. */
///     fseek_err = fseek(stream, -2, SEEK_CUR);
///     if (fseek_err)
///         goto file_error;

/* Search through the image data and read the JPEG markers. */
///     while (!feof(stream)) {

/* Read the JPEG marker and length fields for the sub-section. */
///         if (fread(&marker, sizeof(marker), 1, stream) < 1)
///             break;

///         if (fread(&length, sizeof(length), 1, stream) < 1)
///             break;

/* Convert the marker and length to network order. */
///         marker = LXW_UINT16_NETWORK(marker);
///         length = LXW_UINT16_NETWORK(length);

/* The offset for next fseek() is the field length + type length. */
///         offset = length - 2;

/* Read the height and width in the 0xFFCn elements (except C4, C8 */
/* and CC which aren't SOF markers). */
///         if ((marker & 0xFFF0) == 0xFFC0 && marker != 0xFFC4
///             && marker != 0xFFC8 && marker != 0xFFCC) {
/* Skip 1 byte to height and width. */
///             fseek_err = fseek(stream, 1, SEEK_CUR);
///             if (fseek_err)
///                 goto file_error;

///             if (fread(&height, sizeof(height), 1, stream) < 1)
///                 break;

///             if (fread(&width, sizeof(width), 1, stream) < 1)
///                 break;

///             height = LXW_UINT16_NETWORK(height);
///             width = LXW_UINT16_NETWORK(width);

///             offset -= 9;
///         }

/* Read the DPI in the 0xFFE0 element. */
///         if (marker == 0xFFE0) {
///             uint16_t x_density = 0;
///             uint16_t y_density = 0;
///             uint8_t units = 1;

///             fseek_err = fseek(stream, 7, SEEK_CUR);
///             if (fseek_err)
///                 goto file_error;

///             if (fread(&units, sizeof(units), 1, stream) < 1)
///                 break;

///             if (fread(&x_density, sizeof(x_density), 1, stream) < 1)
///                 break;

///             if (fread(&y_density, sizeof(y_density), 1, stream) < 1)
///                 break;

///             x_density = LXW_UINT16_NETWORK(x_density);
///             y_density = LXW_UINT16_NETWORK(y_density);

///             if (units == 1) {
///                 x_dpi = x_density;
///                 y_dpi = y_density;
///             }

///             if (units == 2) {
///                 x_dpi = x_density * 2.54;
///                 y_dpi = y_density * 2.54;
///             }

///             offset -= 12;
///         }

///         if (marker == 0xFFDA)
///             break;

///         if (!feof(stream)) {
///             fseek_err = fseek(stream, offset, SEEK_CUR);
///             if (fseek_err)
///                 break;
///         }
///     }

/* Ensure that we read some valid data from the file. */
///     if (width == 0)
///         goto file_error;

/* Set the image metadata. */
///     image_props->image_type = LXW_IMAGE_JPEG;
///     image_props->width = width;
///     image_props->height = height;
///     image_props->x_dpi = x_dpi ? x_dpi : 96;
///     image_props->y_dpi = y_dpi ? y_dpi : 96;
///     image_props->extension = lxw_strdup("jpeg");

///     return LXW_NO_ERROR;

/// file_error:
///     LXW_WARN_FORMAT1("worksheet image insertion: "
///                      "no size data found in: %s.", image_props->filename);

///     return LXW_ERROR_IMAGE_DIMENSIONS;
/// }

/// STATIC lxw_error
/// _process_bmp(lxw_object_properties *image_props)
/// {
///     uint32_t width = 0;
///     uint32_t height = 0;
///     double x_dpi = 96;
///     double y_dpi = 96;
///     int fseek_err;

///     FILE *stream = image_props->stream;

/* Skip another 14 bytes to the start of the BMP height/width. */
///     fseek_err = fseek(stream, 14, SEEK_CUR);
///     if (fseek_err)
///         goto file_error;

///     if (fread(&width, sizeof(width), 1, stream) < 1)
///         width = 0;

///     if (fread(&height, sizeof(height), 1, stream) < 1)
///         height = 0;

/* Ensure that we read some valid data from the file. */
///     if (width == 0)
///         goto file_error;

///     height = LXW_UINT32_HOST(height);
///     width = LXW_UINT32_HOST(width);

/* Set the image metadata. */
///     image_props->image_type = LXW_IMAGE_BMP;
///     image_props->width = width;
///     image_props->height = height;
///     image_props->x_dpi = x_dpi;
///     image_props->y_dpi = y_dpi;
///     image_props->extension = lxw_strdup("bmp");

///     return LXW_NO_ERROR;

/// file_error:
///     LXW_WARN_FORMAT1("worksheet image insertion: "
///                      "no size data found in: %s.", image_props->filename);

///     return LXW_ERROR_IMAGE_DIMENSIONS;
/// }

/// STATIC lxw_error
/// _process_gif(lxw_object_properties *image_props)
/// {
///     uint16_t width = 0;
///     uint16_t height = 0;
///     double x_dpi = 96;
///     double y_dpi = 96;
///     int fseek_err;

///     FILE *stream = image_props->stream;

/* Skip another 2 bytes to the start of the GIF height/width. */
///     fseek_err = fseek(stream, 2, SEEK_CUR);
///     if (fseek_err)
///         goto file_error;

///     if (fread(&width, sizeof(width), 1, stream) < 1)
///         width = 0;

///     if (fread(&height, sizeof(height), 1, stream) < 1)
///         height = 0;

/* Ensure that we read some valid data from the file. */
///     if (width == 0)
///         goto file_error;

///     height = LXW_UINT16_HOST(height);
///     width = LXW_UINT16_HOST(width);

/* Set the image metadata. */
///     image_props->image_type = LXW_IMAGE_GIF;
///     image_props->width = width;
///     image_props->height = height;
///     image_props->x_dpi = x_dpi;
///     image_props->y_dpi = y_dpi;
///     image_props->extension = lxw_strdup("gif");

///     return LXW_NO_ERROR;

/// file_error:
///     LXW_WARN_FORMAT1("worksheet image insertion: "
///                      "no size data found in: %s.", image_props->filename);

///     return LXW_ERROR_IMAGE_DIMENSIONS;
/// }

void process_image(object_properties_t& image_props, const std::vector<unsigned char>& buffer)
{
  ///     uint8_t i;
  ///     size_t size_read;
  ///     char buffer[LXW_IMAGE_BUFFER_SIZE];

  if(buffer[1] == 'P' && buffer[2] == 'N' && buffer[3] == 'G')
  {
    process_png(image_props, buffer);
  }
  // TODO Manage other formats
  ///     else if (signature[0] == 0xFF && signature[1] == 0xD8) {
  ///         if (_process_jpeg(image_props) != LXW_NO_ERROR)
  ///             return LXW_ERROR_IMAGE_DIMENSIONS;
  ///     }
  ///     else if (memcmp(signature, "BM", 2) == 0) {
  ///         if (_process_bmp(image_props) != LXW_NO_ERROR)
  ///             return LXW_ERROR_IMAGE_DIMENSIONS;
  ///     }
  ///     else if (memcmp(signature, "GIF8", 4) == 0) {
  ///         if (_process_gif(image_props) != LXW_NO_ERROR)
  ///             return LXW_ERROR_IMAGE_DIMENSIONS;
  ///     }
  ///     else {
  ///         LXW_WARN_FORMAT1("worksheet image insertion: "
  ///                          "unsupported image format for: %s.",
  ///                          image_props->filename);
  ///         return LXW_ERROR_IMAGE_DIMENSIONS;
  ///     }

  // Calculate an MD5 checksum for the image so that we can remove duplicate
  // images to reduce the xlsx file size.
  MD5_CTX md5_context;
  unsigned char md5_checksum[MD5_SIZE];
  MD5_Init(&md5_context);
  MD5_Update(&md5_context, buffer.data(), buffer.size());
  MD5_Final(md5_checksum, &md5_context);
  for(const auto b: md5_checksum)
  {
    image_props.md5_ += std::format("{:02X}", b);
  }
}

// TODO Use dedicated library to get image properties
void get_image_properties(object_properties_t& image_props)
{
  ///     uint8_t i;
  ///     size_t size_read;
  ///     char buffer[LXW_IMAGE_BUFFER_SIZE];

  if(image_props.image_buffer_.empty())
  {
    // Read image.
    std::ifstream image_stream(image_props.filename_, std::ios::binary);
    std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
    process_image(image_props, buffer);
  }
  else
  {
    process_image(image_props, image_props.image_buffer_);
  }
}

/* Conditional formats that refer to the same cell sqref range, like A or
 * B1:B9, need to be written as part of one xml structure. Therefore we need
 * to store them in a RB hash/tree keyed by sqref. Within the RB hash element
 * we then store conditional formats that refer to sqref in a STAILQ list. */
/// lxw_error
/// _store_conditional_format_object(lxw_worksheet *self,
///                                  lxw_cond_format_obj *cond_format)
/// {
///     lxw_cond_format_hash_element tmp_hash_element;
///     lxw_cond_format_hash_element *found_hash_element = NULL;
///     lxw_cond_format_hash_element *new_hash_element = NULL;

/* Create a temp hash element to do the lookup. */
///     LXW_ATTRIBUTE_COPY(tmp_hash_element.sqref, cond_format->sqref);
///     found_hash_element = RB_FIND(lxw_cond_format_hash,
///                                  self->conditional_formats,
///                                  &tmp_hash_element);

///     if (found_hash_element) {
/* If the RB element exists then add the conditional format to the
 * list for the sqref range.*/
///         STAILQ_INSERT_TAIL(found_hash_element->cond_formats, cond_format,
///                            list_pointers);
///     }
///     else {
/* Create a new RB hash element. */
///         new_hash_element = calloc(1, sizeof(lxw_cond_format_hash_element));
///         GOTO_LABEL_ON_MEM_ERROR(new_hash_element, mem_error);

/* Use the sqref as the key. */
///         LXW_ATTRIBUTE_COPY(new_hash_element->sqref, cond_format->sqref);

/* Also create the list where we store the cond format objects. */
///         new_hash_element->cond_formats =
///             calloc(1, sizeof(struct lxw_cond_format_list));
///         GOTO_LABEL_ON_MEM_ERROR(new_hash_element->cond_formats, mem_error);

/* Initialize the list and add the conditional format object. */
///         STAILQ_INIT(new_hash_element->cond_formats);
///         STAILQ_INSERT_TAIL(new_hash_element->cond_formats, cond_format,
///                            list_pointers);

/* Now insert the RB hash element into the tree. */
///         RB_INSERT(lxw_cond_format_hash, self->conditional_formats,
///                   new_hash_element);
///     }

///     return LXW_NO_ERROR;

/// mem_error:
///     free(new_hash_element);
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
/// }

std::string worksheet_t::write_number_cell(std::string_view range, int32_t style_index, const cell_t& cell) const
{
  /// #ifdef USE_DTOA_LIBRARY
  ///     char data[LXW_ATTR_32];

  ///     lxw_sprintf_dbl(data, cell->u.number);

  ///     if (style_index)
  ///         fprintf(self->file,
  ///                 "<c r=\"%s\" s=\"%d\"><v>%s</v></c>",
  ///                 range, style_index, data);
  ///     else
  ///         fprintf(self->file, "<c r=\"%s\"><v>%s</v></c>", range, data);
  /// #else
  if(style_index != 0)
  {
    return std::format(R"(<c r="{}" s="{}"><v>{:.16G}</v></c>)", range, style_index, std::get<double>(cell.data_));
  }
  else
  {
    return std::format(R"(<c r="{}"><v>{:.16G}</v></c>)", range, std::get<double>(cell.data_));
  }
}

std::string worksheet_t::write_string_cell(std::string_view range, int32_t style_index, const cell_t& cell) const
{
  if(style_index != 0)
  {
    return std::format(R"(<c r="{}" s="{}" t="s"><v>{}</v></c>)", range, style_index, std::get<uint32_t>(cell.data_));
  }
  else
  {
    return std::format(R"(<c r="{}" t="s"><v>{}</v></c>)", range, std::get<uint32_t>(cell.data_));
  }
}

/*
 * Write out an inline string. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
/// STATIC void
/// _write_inline_string_cell(lxw_worksheet *self, char *range,
///                           int32_t style_index, cell_t *cell)
/// {
///     char *string = lxw_escape_data(cell->u.string);

/* Add attribute to preserve leading or trailing whitespace. */
///     if (isspace((unsigned char) string[0])
///         || isspace((unsigned char) string[strlen(string) - 1])) {

///         if (style_index)
///             fprintf(self->file,
///                     "<c r=\"%s\" s=\"%d\" t=\"inlineStr\"><is>"
///                     "<t xml:space=\"preserve\">%s</t></is></c>",
///                     range, style_index, string);
///         else
///             fprintf(self->file,
///                     "<c r=\"%s\" t=\"inlineStr\"><is>"
///                     "<t xml:space=\"preserve\">%s</t></is></c>",
///                     range, string);
///     }
///     else {
///         if (style_index)
///             fprintf(self->file,
///                     "<c r=\"%s\" s=\"%d\" t=\"inlineStr\">"
///                     "<is><t>%s</t></is></c>", range, style_index, string);
///         else
///             fprintf(self->file,
///                     "<c r=\"%s\" t=\"inlineStr\">"
///                     "<is><t>%s</t></is></c>", range, string);
///     }

///     free(string);
/// }

/*
 * Write out an inline rich string. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
/// STATIC void
/// _write_inline_rich_string_cell(lxw_worksheet *self, char *range,
///                                int32_t style_index, cell_t *cell)
/// {
///     const char *string = cell->u.string;

///     if (style_index)
///         fprintf(self->file,
///                 "<c r=\"%s\" s=\"%d\" t=\"inlineStr\">"
///                 "<is>%s</is></c>", range, style_index, string);
///     else
///         fprintf(self->file,
///                 "<c r=\"%s\" t=\"inlineStr\">"
///                 "<is>%s</is></c>", range, string);
/// }

/*
 * Write out a formula worksheet cell with a numeric result.
 */
std::string worksheet_t::write_formula_num_cell(const cell_t& cell) const
{
  std::string xml_data = xml_data_element("f", std::get<std::string>(cell.data_));
  xml_data += xml_data_element("v", std::format("{}", cell.formula_result_));

  return xml_data;
}

/*
 * Write out a formula worksheet cell with a numeric result.
 */
std::string worksheet_t::write_formula_str_cell(const cell_t& cell) const
{
  std::string xml_data = xml_data_element("f", std::get<std::string>(cell.data_));
  xml_data += xml_data_element("v", cell.user_data2_);

  return xml_data;
}

/*
 * Write out an array formula worksheet cell with a numeric result.
 */
/// STATIC void
/// _write_array_formula_num_cell(lxw_worksheet *self, cell_t *cell)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char data[LXW_ATTR_32];

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("t", "array");
///     LXW_PUSH_ATTRIBUTES_STR("ref", cell->user_data1);

///     lxw_sprintf_dbl(data, cell->formula_result);

///     lxw_xml_data_element(self->file, "f", cell->u.string, &attributes);
///     lxw_xml_data_element(self->file, "v", data, NULL);

///     LXW_FREE_ATTRIBUTES();
/// }

/*
 * Write out a boolean worksheet cell.
 */
/// STATIC void
/// _write_boolean_cell(lxw_worksheet *self, cell_t *cell)
/// {
///     char data[LXW_ATTR_32];

///     if (cell->u.number == 0.0)
///         data[0] = '0';
///     else
///         data[0] = '1';

///     data[1] = '\0';

///     lxw_xml_data_element(self->file, "v", data, NULL);
/// }

/*
 * Write out a error worksheet cell.
 */
std::string worksheet_t::write_error_cell() const
{
  return xml_data_element("v", "#VALUE!");
}

std::string worksheet_t::write_cell(const cell_t& cell /* TODO , lxw_format *row_format*/) const
{
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  int32_t style_index     = 0;
  const std::string range = rowcol_to_cell(cell.row_num_, cell.col_num_);

  if(cell.format_)
  {
    style_index = get_xf_index_(cell.format_);
  }
  ///     else if (row_format) {
  ///         style_index = lxw_format_get_xf_index(row_format);
  ///     }
  ///     else if (col_num < self->col_formats_max && self->col_formats[col_num]) {
  ///         style_index = lxw_format_get_xf_index(self->col_formats[col_num]);
  ///     }

  if(cell.type_ == cell_types_t::NUMBER_CELL)
  {
    return write_number_cell(range, style_index, cell);
  }

  if(cell.type_ == cell_types_t::STRING_CELL)
  {
    return write_string_cell(range, style_index, cell);
  }

  ///     if (cell->type == INLINE_STRING_CELL) {
  ///         _write_inline_string_cell(self, range, style_index, cell);
  ///         return;
  ///     }

  ///     if (cell->type == INLINE_RICH_STRING_CELL) {
  ///         _write_inline_rich_string_cell(self, range, style_index, cell);
  ///         return;
  ///     }

  // For other cell types use the general functions.
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"r", range}
  };

  if(style_index != 0)
  {
    attributes.emplace_back("s", std::to_string(style_index));
  }

  if(cell.type_ == cell_types_t::FORMULA_CELL)
  {
    // If user_data2 is set then the formula has a string result.
    if(!cell.user_data2_.empty())
    {
      attributes.emplace_back("t", "str");
    }

    std::string xml_data = xml_start_tag("c", attributes);

    if(!cell.user_data2_.empty())
    {
      xml_data += write_formula_str_cell(cell);
    }
    else
    {
      xml_data += write_formula_num_cell(cell);
    }

    xml_data += xml_end_tag("c");

    return xml_data;
  }
  else if(cell.type_ == cell_types_t::BLANK_CELL)
  {
    if(cell.format_)
    {
      return xml_empty_tag("c", attributes);
    }
  }
  ///     else if (cell->type == BOOLEAN_CELL) {
  ///         LXW_PUSH_ATTRIBUTES_STR("t", "b");
  ///         lxw_xml_start_tag(self->file, "c", &attributes);
  ///         _write_boolean_cell(self, cell);
  ///         lxw_xml_end_tag(self->file, "c");
  ///     }
  ///     else if (cell->type == ARRAY_FORMULA_CELL) {
  ///         lxw_xml_start_tag(self->file, "c", &attributes);
  ///         _write_array_formula_num_cell(self, cell);
  ///         lxw_xml_end_tag(self->file, "c");
  ///     }
  ///     else if (cell->type == DYNAMIC_ARRAY_FORMULA_CELL) {
  ///         LXW_PUSH_ATTRIBUTES_STR("cm", "1");
  ///         lxw_xml_start_tag(self->file, "c", &attributes);
  ///         _write_array_formula_num_cell(self, cell);
  ///         lxw_xml_end_tag(self->file, "c");
  ///     }
  else if(cell.type_ == cell_types_t::ERROR_CELL)
  {
    attributes.emplace_back("t", "e");
    attributes.emplace_back("vm", std::format("{}", std::get<uint32_t>(cell.data_)));
    std::string xml_data = xml_start_tag("c", attributes);
    xml_data += write_error_cell();
    xml_data += xml_end_tag("c");

    return xml_data;
  }

  return "";
}

std::string worksheet_t::write_rows() const
{
  std::string xml_data;
  int32_t block_num = -1;
  std::string spans;

  for(auto it = std::begin(table_.rbh_root_); it != std::end(table_.rbh_root_); it++)
  {
    const auto& row = it->second;
    if(row.cells_.empty())
    {
      // Row contains no cells but has height, format or other data.
      // Write a default span for default rows.
      if(default_row_set_)
      {
        xml_data += write_row(row, "1:1");
      }
      else
      {
        xml_data += write_row(row);
      }
    }
    else
    {
      // Row and cell data.
      if(static_cast<int32_t>(row.row_num_ / 16) > block_num)
      {
        spans = calculate_spans(it, std::end(table_.rbh_root_), block_num);
      }

      xml_data += write_row(row, spans);

      if(row.data_changed_)
      {
        for(const auto& [index, cell]: row.cells_)
        {
          xml_data += write_cell(cell /*TODO , row->format*/);
        }
      }

      xml_data += xml_end_tag("row");
    }
  }

  return xml_data;
}

/*
 * Write out the worksheet data as a single row with cells. This method is
 * used when memory optimization is on. A single row is written and the data
 * array is reset. That way only one row of data is kept in memory at any one
 * time. We don't write span data in the optimized case since it is optional.
 */
/// void lxw_worksheet_write_single_row(lxw_worksheet *self)
/// {
///     row_t *row = self->optimize_row;
///     col_num_t col;

/* skip row if it doesn't contain row formatting, cell data or a comment. */
///     if (!(row->row_changed || row->data_changed))
///         return;

/* Write the cells if the row contains data. */
///     if (!row->data_changed) {
/* Row data only. No cells. */
///         _write_row(self, row, NULL);
///     }
///     else {
/* Row and cell data. */
///         _write_row(self, row, NULL);

///         for (col = self->dim_colmin; col <= self->dim_colmax; col++) {
///             if (self->array[col]) {
///                 _write_cell(self, self->array[col], row->format);
///                 _free_cell(self->array[col]);
///                 self->array[col] = NULL;
///             }
///         }

///         lxw_xml_end_tag(self->file, "row");
///     }

/* Reset the row. */
///     row->height = LXW_DEF_ROW_HEIGHT;
///     row->format = NULL;
///     row->hidden = LXW_FALSE;
///     row->level = 0;
///     row->collapsed = LXW_FALSE;
///     row->data_changed = LXW_FALSE;
///     row->row_changed = LXW_FALSE;
/// }

// Set VML image position string based on the header/footer/position.
std::string worksheet_t::get_vml_image_position(image_position_t image_position) const
{
  switch(image_position)
  {
    case image_position_t::HEADER_LEFT:
      return "LH";

    case image_position_t::HEADER_CENTER:
      return "CH";

    case image_position_t::HEADER_RIGHT:
      return "RH";

    case image_position_t::FOOTER_LEFT:
      return "LF";

    case image_position_t::FOOTER_CENTER:
      return "CF";

    case image_position_t::FOOTER_RIGHT:
      return "RF";
  }

  return "";
}

// Process a header/footer image and store it in the correct slot.
void worksheet_t::set_header_footer_image(const std::string& filename, image_position_t image_position)
{
  // Not all slots will have image files.
  if(filename.empty())
  {
    return;
  }

  // Check that the image file exists and can be opened.
  {
    std::ifstream image_stream(filename);
    if(!image_stream)
    {
      throw xwpp_exception_t(std::format("Image file {} doesn't exist or cannot be opened", filename));
    }
  }

  // Create a new object to hold the image properties.
  object_properties_t object_props;

  // Copy other options or set defaults.
  object_props.filename_    = filename;
  // Use the filename as the default description, like Excel.
  // TODO Use basename of file, not full name
  object_props.description_ = filename;

  // Set VML image position string based on the header/footer/position.
  object_props.image_position_ = get_vml_image_position(image_position);

  get_image_properties(object_props);

  header_footer_objs_[static_cast<size_t>(image_position)] = object_props;
  has_header_vml_                                          = true;
}

std::string worksheet_t::write_col_info(const col_options_t& options) const
{
  std::string xml_data;
  std::vector<std::tuple<std::string, std::string>> attributes;

  double width          = options.width_;
  bool has_custom_width = true;
  ///     int32_t xf_index = 0;

  /* Get the format index. */
  ///     if (options->format) {
  ///         xf_index = lxw_format_get_xf_index(options->format);
  ///     }

  // Check if width is the Excel default.
  if(width == DEF_COL_WIDTH)
  {

    // The default col width changes to 0 for hidden columns.
    ///         if (options->hidden)
    ///             width = 0;
    ///         else
    has_custom_width = false;
  }

  attributes.emplace_back("min", std::to_string(options.firstcol_ + 1));
  attributes.emplace_back("max", std::to_string(options.lastcol_ + 1));
  attributes.emplace_back("width", std::format("{}", width));

  ///     if (xf_index)
  ///         LXW_PUSH_ATTRIBUTES_INT("style", xf_index);

  ///     if (options->hidden)
  ///         LXW_PUSH_ATTRIBUTES_STR("hidden", "1");

  if(has_custom_width)
  {
    attributes.emplace_back("customWidth", "1");
  }

  ///     if (options->level)
  ///         LXW_PUSH_ATTRIBUTES_INT("outlineLevel", options->level);

  ///     if (options->collapsed)
  ///         LXW_PUSH_ATTRIBUTES_STR("collapsed", "1");

  xml_data += xml_empty_tag("col", attributes);

  return xml_data;
}

std::string worksheet_t::write_cols() const
{
  if(!col_size_changed_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("cols");

  for(size_t col = 0; col < col_options_.size(); col++)
  {
    if(col_options_[col].firstcol_ == col)
    {
      xml_data += write_col_info(col_options_[col]);
    }
  }

  xml_data += xml_end_tag("cols");

  return xml_data;
}

std::string worksheet_t::write_merge_cell(const merged_range_t& merged_range) const
{
  // Convert the merge dimensions to a cell range.
  std::string ref =
      rowcol_to_range(merged_range.first_row_, merged_range.first_col_, merged_range.last_row_, merged_range.last_col_);

  return xml_empty_tag("mergeCell", {
                                        {"ref", ref}
  });
}

std::string worksheet_t::write_merge_cells() const
{
  if(merged_ranges_.empty())
  {
    return "";
  }

  std::string xml_data = xml_start_tag("mergeCells", {
                                                         {"count", std::to_string(merged_ranges_.size())}
  });
  for(const auto& merged_range: merged_ranges_)
  {
    xml_data += write_merge_cell(merged_range);
  }
  xml_data += xml_end_tag("mergeCells");

  return xml_data;
}

std::string worksheet_t::write_odd_header() const
{
  return xml_data_element("oddHeader", header_);
}

std::string worksheet_t::write_odd_footer() const
{
  return xml_data_element("oddFooter", footer_);
}

std::string worksheet_t::write_header_footer() const
{
  if(!header_footer_changed_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("headerFooter");
  if(!header_.empty())
  {
    xml_data += write_odd_header();
  }
  if(!footer_.empty())
  {
    xml_data += write_odd_footer();
  }
  xml_data += xml_end_tag("headerFooter");

  return xml_data;
}

/// STATIC void _worksheet_write_page_set_up_pr(lxw_worksheet *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     if (!self->fit_page)
///         return;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("fitToPage", "1");

///     lxw_xml_empty_tag(self->file, "pageSetUpPr", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_tab_color() const
{
  if(tab_color_ == color_t::UNSET)
  {
    return "";
  }

  return xml_empty_tag("tabColor",
                       {
                           {"rgb", std::format("FF{:06X}", static_cast<uint32_t>(tab_color_) & COLOR_MASK)}
  });
}

/// STATIC void _worksheet_write_outline_pr(lxw_worksheet *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     if (!self->outline_changed)
///         return;

///     LXW_INIT_ATTRIBUTES();

///     if (self->outline_style)
///         LXW_PUSH_ATTRIBUTES_STR("applyStyles", "1");

///     if (!self->outline_below)
///         LXW_PUSH_ATTRIBUTES_STR("summaryBelow", "0");

///     if (!self->outline_right)
///         LXW_PUSH_ATTRIBUTES_STR("summaryRight", "0");

///     if (!self->outline_on)
///         LXW_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

///     lxw_xml_empty_tag(self->file, "outlinePr", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_sheet_pr() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(!fit_page_ && !filter_on_ && tab_color_ == color_t::UNSET && !outline_changed_
     ///         && !self->vba_codename && !self->is_chartsheet
  )
  {
    return "";
  }

  ///     if (self->vba_codename)
  ///         LXW_PUSH_ATTRIBUTES_STR("codeName", self->vba_codename);

  ///     if (self->filter_on)
  ///         LXW_PUSH_ATTRIBUTES_STR("filterMode", "1");

  if(fit_page_ || tab_color_ != color_t::UNSET || outline_changed_)
  {
    std::string xml_data = xml_start_tag("sheetPr", attributes);
    xml_data += write_tab_color();
    ///         _worksheet_write_outline_pr(self);
    ///         _worksheet_write_page_set_up_pr(self);
    xml_data += xml_end_tag("sheetPr");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("sheetPr", attributes);
  }
}

std::string worksheet_t::write_brk(uint32_t id, uint32_t max) const
{
  return xml_empty_tag("brk", {
                                  {"id",  std::to_string(id) },
                                  {"max", std::to_string(max)},
                                  {"man", "1"                },
  });
}

std::string worksheet_t::write_row_breaks() const
{
  if(hbreaks_.empty())
  {
    return "";
  }

  std::string xml_data = xml_start_tag("rowBreaks", {
                                                        {"count",            std::to_string(hbreaks_.size())},
                                                        {"manualBreakCount", std::to_string(hbreaks_.size())},
  });

  for(const auto row: hbreaks_)
  {
    xml_data += write_brk(row, COL_MAX - 1);
  }
  xml_data += xml_end_tag("rowBreaks");

  return xml_data;
}

std::string worksheet_t::write_col_breaks() const
{
  if(vbreaks_.empty())
  {
    return "";
  }

  std::string xml_data = xml_start_tag("colBreaks", {
                                                        {"count",            std::to_string(vbreaks_.size())},
                                                        {"manualBreakCount", std::to_string(vbreaks_.size())},
  });

  for(const auto col: vbreaks_)
  {
    xml_data += write_brk(col, ROW_MAX - 1);
  }
  xml_data += xml_end_tag("colBreaks");

  return xml_data;
}

/// STATIC void _worksheet_write_filter(lxw_worksheet *self, const char *str, double num,
///                         uint8_t criteria)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     if (criteria == LXW_FILTER_CRITERIA_BLANKS)
///         return;

///     LXW_INIT_ATTRIBUTES();

///     if (str)
///         LXW_PUSH_ATTRIBUTES_STR("val", str);
///     else
///         LXW_PUSH_ATTRIBUTES_DBL("val", num);

///     lxw_xml_empty_tag(self->file, "filter", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void _worksheet_write_filter_standard(lxw_worksheet *self,
///                                  lxw_filter_rule_obj *filter)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (filter->has_blanks) {
///         LXW_PUSH_ATTRIBUTES_STR("blank", "1");
///     }

///     if (filter->type == LXW_FILTER_TYPE_SINGLE && filter->has_blanks) {
///         lxw_xml_empty_tag(self->file, "filters", &attributes);
///     }
///     else {
///         lxw_xml_start_tag(self->file, "filters", &attributes);

/* Write the filter element. */
///         if (filter->type == LXW_FILTER_TYPE_SINGLE) {
///             _worksheet_write_filter(self, filter->value1_string,
///                                     filter->value1, filter->criteria1);
///         }
///         else if (filter->type == LXW_FILTER_TYPE_AND
///                  || filter->type == LXW_FILTER_TYPE_OR) {
///             _worksheet_write_filter(self, filter->value1_string,
///                                     filter->value1, filter->criteria1);
///             _worksheet_write_filter(self, filter->value2_string,
///                                     filter->value2, filter->criteria2);
///         }

///         lxw_xml_end_tag(self->file, "filters");
///     }

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void _worksheet_write_custom_filter(lxw_worksheet *self, const char *str,
///                                double num, uint8_t criteria)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (criteria == LXW_FILTER_CRITERIA_NOT_EQUAL_TO)
///         LXW_PUSH_ATTRIBUTES_STR("operator", "notEqual");
///     if (criteria == LXW_FILTER_CRITERIA_GREATER_THAN)
///         LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
///     else if (criteria == LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO)
///         LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
///     else if (criteria == LXW_FILTER_CRITERIA_LESS_THAN)
///         LXW_PUSH_ATTRIBUTES_STR("operator", "lessThan");
///     else if (criteria == LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO)
///         LXW_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");

///     if (str)
///         LXW_PUSH_ATTRIBUTES_STR("val", str);
///     else
///         LXW_PUSH_ATTRIBUTES_DBL("val", num);

///     lxw_xml_empty_tag(self->file, "customFilter", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void _worksheet_write_filter_list(lxw_worksheet *self, lxw_filter_rule_obj *filter)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     uint16_t i;

///     LXW_INIT_ATTRIBUTES();

///     if (filter->has_blanks) {
///         LXW_PUSH_ATTRIBUTES_STR("blank", "1");
///     }

///     lxw_xml_start_tag(self->file, "filters", &attributes);

///     for (i = 0; i < filter->num_list_filters; i++) {
///         _worksheet_write_filter(self, filter->list[i], 0, 0);
///     }

///     lxw_xml_end_tag(self->file, "filters");

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_filter_custom(lxw_worksheet *self,
///                                lxw_filter_rule_obj *filter)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (filter->type == LXW_FILTER_TYPE_AND)
///         LXW_PUSH_ATTRIBUTES_STR("and", "1");

///     lxw_xml_start_tag(self->file, "customFilters", &attributes);

/* Write the filter element. */
///     if (filter->type == LXW_FILTER_TYPE_SINGLE) {
///         _worksheet_write_custom_filter(self, filter->value1_string,
///                                        filter->value1, filter->criteria1);
///     }
///     else if (filter->type == LXW_FILTER_TYPE_AND
///              || filter->type == LXW_FILTER_TYPE_OR) {
///         _worksheet_write_custom_filter(self, filter->value1_string,
///                                        filter->value1, filter->criteria1);
///         _worksheet_write_custom_filter(self, filter->value2_string,
///                                        filter->value2, filter->criteria2);
///     }
///
///     lxw_xml_end_tag(self->file, "customFilters");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_filter_column(lxw_worksheet *self,
///                                lxw_filter_rule_obj *filter)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     if (!filter)
///         return;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("colId", filter->col_num);
///
///     lxw_xml_start_tag(self->file, "filterColumn", &attributes);
///
///     if (filter->list)
///         _worksheet_write_filter_list(self, filter);
///     else if (filter->is_custom)
///         _worksheet_write_filter_custom(self, filter);
///     else
///         _worksheet_write_filter_standard(self, filter);
///
///     lxw_xml_end_tag(self->file, "filterColumn");
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_auto_filter() const
{
  const std::vector<std::tuple<std::string, std::string>> attributes;
  std::string xml_data;
  ///     char range[LXW_MAX_CELL_RANGE_LENGTH];
  ///     uint16_t i;

  ///     if (!self->autofilter.in_use)
  ///         return;

  ///     lxw_rowcol_to_range(range,
  ///                         self->autofilter.first_row,
  ///                         self->autofilter.first_col,
  ///                         self->autofilter.last_row,
  ///                         self->autofilter.last_col);

  ///     LXW_INIT_ATTRIBUTES();
  ///     LXW_PUSH_ATTRIBUTES_STR("ref", range);

  ///     if (self->autofilter.has_rules) {
  ///         lxw_xml_start_tag(self->file, "autoFilter", &attributes);

  ///         for (i = 0; i < self->num_filter_rules; i++)
  ///             _worksheet_write_filter_column(self, self->filter_rules[i]);

  ///         lxw_xml_end_tag(self->file, "autoFilter");

  ///     }
  ///     else {
  ///         lxw_xml_empty_tag(self->file, "autoFilter", &attributes);
  ///     }

  return xml_data;
}

std::string worksheet_t::write_hyperlink_external(row_num_t row_num, col_num_t col_num, const std::string& location,
                                                  const std::string& tooltip, uint16_t id) const
{
  const std::string range = rowcol_to_cell(row_num, col_num);
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"ref", range},
      {"r:id", std::format("rId{}", id)}
  };

  if(!location.empty())
  {
    attributes.emplace_back("location", location);
  }

  if(!tooltip.empty())
  {
    attributes.emplace_back("tooltip", tooltip);
  }

  return xml_empty_tag("hyperlink", attributes);
}

std::string worksheet_t::write_hyperlink_internal(row_num_t row_num, col_num_t col_num, const std::string& location,
                                                  const std::string& display, const std::string& tooltip) const
{
  const std::string range = rowcol_to_cell(row_num, col_num);
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"ref", range}
  };

  if(!location.empty())
  {
    attributes.emplace_back("location", location);
  }

  if(!tooltip.empty())
  {
    attributes.emplace_back("tooltip", tooltip);
  }

  if(!display.empty())
  {
    attributes.emplace_back("display", display);
  }

  return xml_empty_tag("hyperlink", attributes);
}

// TODO Add again const (remove to allow increment of rel_count_)
std::string worksheet_t::write_hyperlinks()
{
  ///     row_t *row;
  ///     cell_t *link;
  ///     lxw_rel_tuple *relationship;

  if(hyperlinks_.rbh_root_.empty())
  {
    return "";
  }

  // Write the hyperlink elements.
  std::string xml_data = xml_start_tag("hyperlinks");
  for(const auto& [index, row]: hyperlinks_.rbh_root_)
  {
    for(const auto& [index, link]: row.cells_)
    {
      if(link.type_ == cell_types_t::HYPERLINK_URL || link.type_ == cell_types_t::HYPERLINK_EXTERNAL)
      {
        rel_count_++;
        external_hyperlinks_.emplace_back("/hyperlink", std::get<std::string>(link.data_), "External");
        xml_data +=
            write_hyperlink_external(link.row_num_, link.col_num_, link.user_data1_, link.user_data2_, rel_count_);
      }

      if(link.type_ == cell_types_t::HYPERLINK_INTERNAL)
      {
        xml_data += write_hyperlink_internal(link.row_num_, link.col_num_, std::get<std::string>(link.data_),
                                             link.user_data1_, link.user_data2_);
      }
    }
  }

  xml_data += xml_end_tag("hyperlinks");

  return xml_data;
}

std::string worksheet_t::write_sheet_protection() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(!protection_.is_configured_)
  {
    return "";
  }

  if(!protection_.hash_.empty())
  {
    attributes.emplace_back("password", protection_.hash_);
  }

  if(!protection_.no_sheet_)
  {
    attributes.emplace_back("sheet", "1");
  }

  if(!protection_.no_content_)
  {
    attributes.emplace_back("content", "1");
  }

  if(!protection_.objects_)
  {
    attributes.emplace_back("objects", "1");
  }

  if(!protection_.scenarios_)
  {
    attributes.emplace_back("scenarios", "1");
  }

  if(protection_.format_cells_)
  {
    attributes.emplace_back("formatCells", "0");
  }

  if(protection_.format_columns_)
  {
    attributes.emplace_back("formatColumns", "0");
  }

  if(protection_.format_rows_)
  {
    attributes.emplace_back("formatRows", "0");
  }

  if(protection_.insert_columns_)
  {
    attributes.emplace_back("insertColumns", "0");
  }

  if(protection_.insert_rows_)
  {
    attributes.emplace_back("insertRows", "0");
  }

  if(protection_.insert_hyperlinks_)
  {
    attributes.emplace_back("insertHyperlinks", "0");
  }

  if(protection_.delete_columns_)
  {
    attributes.emplace_back("deleteColumns", "0");
  }

  if(protection_.delete_rows_)
  {
    attributes.emplace_back("deleteRows", "0");
  }

  if(protection_.no_select_locked_cells_)
  {
    attributes.emplace_back("selectLockedCells", "1");
  }

  if(protection_.sort_)
  {
    attributes.emplace_back("sort", "0");
  }

  if(protection_.autofilter_)
  {
    attributes.emplace_back("autoFilter", "0");
  }

  if(protection_.pivot_tables_)
  {
    attributes.emplace_back("pivotTables", "0");
  }

  if(protection_.no_select_unlocked_cells_)
  {
    attributes.emplace_back("selectUnlockedCells", "1");
  }

  return xml_empty_tag("sheetProtection", attributes);
}

std::string worksheet_t::write_legacy_drawing()
{
  if(!has_vml_)
  {
    return "";
  }

  rel_count_++;

  return xml_empty_tag("legacyDrawing", {
                                            {"r:id", std::format("rId{}", rel_count_)}
  });
}

std::string worksheet_t::write_legacy_drawing_hf()
{
  if(!has_header_vml_)
  {
    return "";
  }
  else
  {
    rel_count_++;
  }

  return xml_empty_tag("legacyDrawingHF", {
                                              {"r:id", std::format("rId{}", rel_count_)}
  });
}

std::string worksheet_t::write_picture() const
{
  std::string xml_data;
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

  ///     if (!self->has_background_image)
  ///         return;
  ///     else
  ///         self->rel_count++;

  ///     lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", self->rel_count);
  ///     LXW_INIT_ATTRIBUTES();
  ///     LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

  ///     lxw_xml_empty_tag(self->file, "picture", &attributes);

  return xml_data;
}

std::string worksheet_t::write_drawing(uint16_t id) const
{
  return xml_empty_tag("drawing", {
                                      {"r:id", std::format("rId{}", id)}
  });
}

std::string worksheet_t::write_drawings()
{
  if(!drawing_)
  {
    return "";
  }

  rel_count_++;
  return write_drawing(rel_count_);
}

/// STATIC void
/// _worksheet_write_formula1_num(lxw_worksheet *self, double number)
/// {
///     char data[LXW_ATTR_32];
///
///     lxw_sprintf_dbl(data, number);
///
///     lxw_xml_data_element(self->file, "formula1", data, NULL);
/// }

/// STATIC void
/// _worksheet_write_formula1_str(lxw_worksheet *self, char *str)
/// {
///     lxw_xml_data_element(self->file, "formula1", str, NULL);
/// }

/// STATIC void
/// _worksheet_write_formula2_num(lxw_worksheet *self, double number)
/// {
///     char data[LXW_ATTR_32];
///
///     lxw_sprintf_dbl(data, number);
///
///     lxw_xml_data_element(self->file, "formula2", data, NULL);
/// }

/// STATIC void
/// _worksheet_write_formula2_str(lxw_worksheet *self, char *str)
/// {
///     lxw_xml_data_element(self->file, "formula2", str, NULL);
/// }

/// STATIC void
/// _worksheet_write_data_validation(lxw_worksheet *self,
///                                  lxw_data_val_obj *validation)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     uint8_t is_between = 0;
///
///     LXW_INIT_ATTRIBUTES();
///
///     switch (validation->validate) {
///         case LXW_VALIDATION_TYPE_INTEGER:
///         case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
///             LXW_PUSH_ATTRIBUTES_STR("type", "whole");
///             break;
///         case LXW_VALIDATION_TYPE_DECIMAL:
///         case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
///             LXW_PUSH_ATTRIBUTES_STR("type", "decimal");
///             break;
///         case LXW_VALIDATION_TYPE_LIST:
///         case LXW_VALIDATION_TYPE_LIST_FORMULA:
///             LXW_PUSH_ATTRIBUTES_STR("type", "list");
///             break;
///         case LXW_VALIDATION_TYPE_DATE:
///         case LXW_VALIDATION_TYPE_DATE_FORMULA:
///         case LXW_VALIDATION_TYPE_DATE_NUMBER:
///             LXW_PUSH_ATTRIBUTES_STR("type", "date");
///             break;
///         case LXW_VALIDATION_TYPE_TIME:
///         case LXW_VALIDATION_TYPE_TIME_FORMULA:
///         case LXW_VALIDATION_TYPE_TIME_NUMBER:
///             LXW_PUSH_ATTRIBUTES_STR("type", "time");
///             break;
///         case LXW_VALIDATION_TYPE_LENGTH:
///         case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
///             LXW_PUSH_ATTRIBUTES_STR("type", "textLength");
///             break;
///         case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
///             LXW_PUSH_ATTRIBUTES_STR("type", "custom");
///             break;
///     }

///     switch (validation->criteria) {
///         case LXW_VALIDATION_CRITERIA_EQUAL_TO:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "equal");
///             break;
///         case LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "notEqual");
///             break;
///         case LXW_VALIDATION_CRITERIA_LESS_THAN:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "lessThan");
///             break;
///         case LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");
///             break;
///         case LXW_VALIDATION_CRITERIA_GREATER_THAN:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
///             break;
///         case LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
///             LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
///             break;
///         case LXW_VALIDATION_CRITERIA_BETWEEN:
///             /* Between is the default for 2 formulas and isn't added. */
///             is_between = 1;
///             break;
///         case LXW_VALIDATION_CRITERIA_NOT_BETWEEN:
///             is_between = 1;
///             LXW_PUSH_ATTRIBUTES_STR("operator", "notBetween");
///             break;
///     }

///     if (validation->error_type == LXW_VALIDATION_ERROR_TYPE_WARNING)
///         LXW_PUSH_ATTRIBUTES_STR("errorStyle", "warning");
///
///     if (validation->error_type == LXW_VALIDATION_ERROR_TYPE_INFORMATION)
///         LXW_PUSH_ATTRIBUTES_STR("errorStyle", "information");
///
///     if (validation->ignore_blank)
///         LXW_PUSH_ATTRIBUTES_INT("allowBlank", 1);
///
///     if (validation->dropdown == LXW_VALIDATION_OFF)
///         LXW_PUSH_ATTRIBUTES_INT("showDropDown", 1);
///
///     if (validation->show_input)
///         LXW_PUSH_ATTRIBUTES_INT("showInputMessage", 1);
///
///     if (validation->show_error)
///         LXW_PUSH_ATTRIBUTES_INT("showErrorMessage", 1);
///
///     if (validation->error_title)
///         LXW_PUSH_ATTRIBUTES_STR("errorTitle", validation->error_title);
///
///     if (validation->error_message)
///         LXW_PUSH_ATTRIBUTES_STR("error", validation->error_message);
///
///     if (validation->input_title)
///         LXW_PUSH_ATTRIBUTES_STR("promptTitle", validation->input_title);
///
///     if (validation->input_message)
///         LXW_PUSH_ATTRIBUTES_STR("prompt", validation->input_message);
///
///     LXW_PUSH_ATTRIBUTES_STR("sqref", validation->sqref);
///
///     if (validation->validate == LXW_VALIDATION_TYPE_ANY)
///         lxw_xml_empty_tag(self->file, "dataValidation", &attributes);
///     else
///         lxw_xml_start_tag(self->file, "dataValidation", &attributes);
///
///     /* Write the formula1 and formula2 elements. */
///     switch (validation->validate) {
///         case LXW_VALIDATION_TYPE_INTEGER:
///         case LXW_VALIDATION_TYPE_DECIMAL:
///         case LXW_VALIDATION_TYPE_LENGTH:
///         case LXW_VALIDATION_TYPE_DATE:
///         case LXW_VALIDATION_TYPE_TIME:
///         case LXW_VALIDATION_TYPE_DATE_NUMBER:
///         case LXW_VALIDATION_TYPE_TIME_NUMBER:
///             _worksheet_write_formula1_num(self, validation->value_number);
///             if (is_between)
///                 _worksheet_write_formula2_num(self,
///                                               validation->maximum_number);
///             break;
///         case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
///         case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
///         case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
///         case LXW_VALIDATION_TYPE_DATE_FORMULA:
///         case LXW_VALIDATION_TYPE_TIME_FORMULA:
///         case LXW_VALIDATION_TYPE_LIST:
///         case LXW_VALIDATION_TYPE_LIST_FORMULA:
///         case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
///             _worksheet_write_formula1_str(self, validation->value_formula);
///             if (is_between)
///                 _worksheet_write_formula2_str(self,
///                                               validation->maximum_formula);
///             break;
///     }
///
///     if (validation->validate != LXW_VALIDATION_TYPE_ANY)
///         lxw_xml_end_tag(self->file, "dataValidation");
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_data_validations() const
{
  std::string xml_data;
  ///    struct xml_attribute_list attributes;
  ///    struct xml_attribute *attribute;
  ///    lxw_data_val_obj *data_validation;

  ///    if (self->num_validations == 0)
  ///        return;

  ///     LXW_PUSH_ATTRIBUTES_INT("count", self->num_validations);

  ///     lxw_xml_start_tag(self->file, "dataValidations", &attributes);

  ///     STAILQ_FOREACH(data_validation, self->data_validations, list_pointers)
  ///     {
  /* Write the dataValidation element. */
  ///         _worksheet_write_data_validation(self, data_validation);
  ///     }

  ///     lxw_xml_end_tag(self->file, "dataValidations");

  return xml_data;
}

/// STATIC void
/// _worksheet_write_formula_str(lxw_worksheet *self, char *data)
/// {
///     lxw_xml_data_element(self->file, "formula", data, NULL);
/// }

/// STATIC void
/// _worksheet_write_formula_num(lxw_worksheet *self, double num)
/// {
///     char data[LXW_ATTR_32];
///
///     lxw_sprintf_dbl(data, num);
///     lxw_xml_data_element(self->file, "formula", data, NULL);
/// }

/// STATIC void
/// _worksheet_write_ext(lxw_worksheet *self, char *uri)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char xmlns_x_14[] =
///         "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("xmlns:x14", xmlns_x_14);
///     LXW_PUSH_ATTRIBUTES_STR("uri", uri);
///
///     lxw_xml_start_tag(self->file, "ext", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_data_bar_ext(lxw_worksheet *self,
///                               lxw_cond_format_obj *cond_format)
/// {
///     /* Create a pseudo GUID for each unique Excel 2010 data bar. */
///     cond_format->guid = calloc(1, LXW_GUID_LENGTH);
///     lxw_snprintf(cond_format->guid, LXW_GUID_LENGTH,
///                  "{DA7ABA51-AAAA-BBBB-%04X-%012X}",
///                  self->index + 1, ++self->data_bar_2010_index);
///
///     lxw_xml_start_tag(self->file, "extLst", NULL);
///
///     _worksheet_write_ext(self, "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}");
///
///     lxw_xml_data_element(self->file, "x14:id", cond_format->guid, NULL);
///
///     lxw_xml_end_tag(self->file, "ext");
///     lxw_xml_end_tag(self->file, "extLst");
/// }

/// STATIC void
/// _worksheet_write_color(lxw_worksheet *self, lxw_color_t color)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb[LXW_ATTR_32];
///
///     lxw_snprintf(rgb, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("rgb", rgb);
///
///     lxw_xml_empty_tag(self->file, "color", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cfvo_str(lxw_worksheet *self, uint8_t rule_type,
///                           char *value, uint8_t data_bar_2010)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM)
///         LXW_PUSH_ATTRIBUTES_STR("type", "min");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER)
///         LXW_PUSH_ATTRIBUTES_STR("type", "num");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT)
///         LXW_PUSH_ATTRIBUTES_STR("type", "percent");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE)
///         LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA)
///         LXW_PUSH_ATTRIBUTES_STR("type", "formula");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM)
///         LXW_PUSH_ATTRIBUTES_STR("type", "max");
///
///     if (!data_bar_2010 || (rule_type != LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///                            && rule_type != LXW_CONDITIONAL_RULE_TYPE_MAXIMUM))
///         LXW_PUSH_ATTRIBUTES_STR("val", value);
///
///     lxw_xml_empty_tag(self->file, "cfvo", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cfvo_num(lxw_worksheet *self, uint8_t rule_type,
///                           double value, uint8_t data_bar_2010)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM)
///         LXW_PUSH_ATTRIBUTES_STR("type", "min");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER)
///         LXW_PUSH_ATTRIBUTES_STR("type", "num");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT)
///         LXW_PUSH_ATTRIBUTES_STR("type", "percent");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE)
///         LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA)
///         LXW_PUSH_ATTRIBUTES_STR("type", "formula");
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM)
///         LXW_PUSH_ATTRIBUTES_STR("type", "max");
///
///     if (!data_bar_2010 || (rule_type != LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///                            && rule_type != LXW_CONDITIONAL_RULE_TYPE_MAXIMUM))
///         LXW_PUSH_ATTRIBUTES_DBL("val", value);
///
///     lxw_xml_empty_tag(self->file, "cfvo", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_icon_set(lxw_worksheet *self,
///                           lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char *icon_set[] = {
///         "3Arrows",
///         "3ArrowsGray",
///         "3Flags",
///         "3TrafficLights",
///         "3TrafficLights2",
///         "3Signs",
///         "3Symbols",
///         "3Symbols2",
///         "4Arrows",
///         "4ArrowsGray",
///         "4RedToBlack",
///         "4Rating",
///         "4TrafficLights",
///         "5Arrows",
///         "5ArrowsGray",
///         "5Rating",
///         "5Quarters",
///     };
///     uint8_t percent = LXW_CONDITIONAL_RULE_TYPE_PERCENT;
///     uint8_t style = cond_format->icon_style;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (style != LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED)
///         LXW_PUSH_ATTRIBUTES_STR("iconSet", icon_set[style]);
///
///     if (cond_format->reverse_icons == LXW_TRUE)
///         LXW_PUSH_ATTRIBUTES_STR("reverse", "1");
///
///     if (cond_format->icons_only == LXW_TRUE)
///         LXW_PUSH_ATTRIBUTES_STR("showValue", "0");
///
///     lxw_xml_start_tag(self->file, "iconSet", &attributes);
///
///     if (style < LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED) {
///         _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 33, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 67, LXW_FALSE);
///     }
///
///     if (style >= LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED
///         && style < LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED) {
///         _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 25, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 50, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 75, LXW_FALSE);
///     }
///
///     if (style >= LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED
///         && style <= LXW_CONDITIONAL_ICONS_5_QUARTERS) {
///         _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 20, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 40, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 60, LXW_FALSE);
///         _worksheet_write_cfvo_num(self, percent, 80, LXW_FALSE);
///     }
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_icons(lxw_worksheet *self,
///                                lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     _worksheet_write_icon_set(self, cond_format);
///
///     lxw_xml_end_tag(self->file, "iconSet");
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_data_bar(lxw_worksheet *self,
///                           lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (cond_format->bar_only)
///         LXW_PUSH_ATTRIBUTES_STR("showValue", "0");
///
///     lxw_xml_start_tag(self->file, "dataBar", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_data_bar(lxw_worksheet *self,
///                                   lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     _worksheet_write_data_bar(self, cond_format);
///
///     if (cond_format->min_value_string) {
///         _worksheet_write_cfvo_str(self, cond_format->min_rule_type,
///                                   cond_format->min_value_string,
///                                   cond_format->data_bar_2010);
///     }
///     else {
///         _worksheet_write_cfvo_num(self, cond_format->min_rule_type,
///                                   cond_format->min_value,
///                                   cond_format->data_bar_2010);
///     }
///
///     if (cond_format->max_value_string) {
///         _worksheet_write_cfvo_str(self, cond_format->max_rule_type,
///                                   cond_format->max_value_string,
///                                   cond_format->data_bar_2010);
///     }
///     else {
///         _worksheet_write_cfvo_num(self, cond_format->max_rule_type,
///                                   cond_format->max_value,
///                                   cond_format->data_bar_2010);
///     }
///
///     _worksheet_write_color(self, cond_format->bar_color);
///
///     lxw_xml_end_tag(self->file, "dataBar");
///
///     if (cond_format->data_bar_2010)
///         _worksheet_write_data_bar_ext(self, cond_format);
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_color_scale(lxw_worksheet *self,
///                                      lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     lxw_xml_start_tag(self->file, "colorScale", NULL);
///
///     if (cond_format->min_value_string) {
///         _worksheet_write_cfvo_str(self, cond_format->min_rule_type,
///                                   cond_format->min_value_string, LXW_FALSE);
///     }
///     else {
///         _worksheet_write_cfvo_num(self, cond_format->min_rule_type,
///                                   cond_format->min_value, LXW_FALSE);
///     }
///
///     if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
///         if (cond_format->mid_value_string) {
///             _worksheet_write_cfvo_str(self, cond_format->mid_rule_type,
///                                       cond_format->mid_value_string,
///                                       LXW_FALSE);
///         }
///         else {
///             _worksheet_write_cfvo_num(self, cond_format->mid_rule_type,
///                                       cond_format->mid_value, LXW_FALSE);
///         }
///     }
///
///     if (cond_format->max_value_string) {
///         _worksheet_write_cfvo_str(self, cond_format->max_rule_type,
///                                   cond_format->max_value_string, LXW_FALSE);
///     }
///     else {
///         _worksheet_write_cfvo_num(self, cond_format->max_rule_type,
///                                   cond_format->max_value, LXW_FALSE);
///     }
///
///     _worksheet_write_color(self, cond_format->min_color);
///
///     if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE)
///         _worksheet_write_color(self, cond_format->mid_color);
///
///     _worksheet_write_color(self, cond_format->max_color);
///
///     lxw_xml_end_tag(self->file, "colorScale");
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_formula(lxw_worksheet *self,
///                                  lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     _worksheet_write_formula_str(self, cond_format->min_value_string);
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_top(lxw_worksheet *self,
///                              lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     if (cond_format->criteria ==
///         LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT)
///         LXW_PUSH_ATTRIBUTES_INT("percent", 1);
///
///     if (cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM)
///         LXW_PUSH_ATTRIBUTES_INT("bottom", 1);
///
///     /* Rank must be an int in the range 1-1000 . */
///     if (cond_format->min_value < 1.0 || cond_format->min_value > 1000.0)
///         LXW_PUSH_ATTRIBUTES_DBL("rank", 10);
///     else
///         LXW_PUSH_ATTRIBUTES_DBL("rank", (uint16_t) cond_format->min_value);
///
///     lxw_xml_empty_tag(self->file, "cfRule", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_duplicate(lxw_worksheet *self,
///                                    lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     /* Set the attributes common to all rule types. */
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     lxw_xml_empty_tag(self->file, "cfRule", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_average(lxw_worksheet *self,
///                                  lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     uint8_t criteria = cond_format->criteria;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
///         LXW_PUSH_ATTRIBUTES_INT("aboveAverage", 0);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL)
///         LXW_PUSH_ATTRIBUTES_INT("equalAverage", 1);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW)
///         LXW_PUSH_ATTRIBUTES_INT("stdDev", 1);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW)
///         LXW_PUSH_ATTRIBUTES_INT("stdDev", 2);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE
///         || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
///         LXW_PUSH_ATTRIBUTES_INT("stdDev", 3);
///
///     lxw_xml_empty_tag(self->file, "cfRule", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_time_period(lxw_worksheet *self,
///                                      lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char formula[LXW_MAX_ATTRIBUTE_LENGTH];
///     uint8_t pos;
///     uint8_t criteria = cond_format->criteria;
///     char *first_cell = cond_format->first_cell;
///     char *time_periods[] = {
///         "yesterday",
///         "today",
///         "tomorrow",
///         "last7Days",
///         "lastWeek",
///         "thisWeek",
///         "nextWeek",
///         "lastMonth",
///         "thisMonth",
///         "nextMonth",
///     };
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     pos = criteria - LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY;
///     LXW_PUSH_ATTRIBUTES_STR("timePeriod", time_periods[pos]);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "FLOOR(%s,1)=TODAY()-1", first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "FLOOR(%s,1)=TODAY()", first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "FLOOR(%s,1)=TODAY()+1", first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())",
///                      first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),"
///                      "TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))",
///                      first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,"
///                      "ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))",
///                      first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),"
///                      "ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))",
///                      first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(MONTH(%s)=MONTH(TODAY())-1,OR(YEAR(%s)=YEAR("
///                      "TODAY()),AND(MONTH(%s)=1,YEAR(A1)=YEAR(TODAY())-1)))",
///                      first_cell, first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))",
///                      first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {
///         lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
///                      "AND(MONTH(%s)=MONTH(TODAY())+1,OR(YEAR(%s)=YEAR("
///                      "TODAY()),AND(MONTH(%s)=12,YEAR(%s)=YEAR(TODAY())+1)))",
///                      first_cell, first_cell, first_cell, first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_blanks(lxw_worksheet *self,
///                                 lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char formula[LXW_ATTR_32];
///     uint8_t type = cond_format->type;
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     if (type == LXW_CONDITIONAL_TYPE_BLANKS) {
///         lxw_snprintf(formula, LXW_ATTR_32, "LEN(TRIM(%s))=0",
///                      cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (type == LXW_CONDITIONAL_TYPE_NO_BLANKS) {
///         lxw_snprintf(formula, LXW_ATTR_32, "LEN(TRIM(%s))>0",
///                      cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (type == LXW_CONDITIONAL_TYPE_ERRORS) {
///         lxw_snprintf(formula, LXW_ATTR_32, "ISERROR(%s)",
///                      cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (type == LXW_CONDITIONAL_TYPE_NO_ERRORS) {
///         lxw_snprintf(formula, LXW_ATTR_32, "NOT(ISERROR(%s))",
///                      cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_text(lxw_worksheet *self,
///                               lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     uint8_t pos;
///     char formula[LXW_ATTR_32 * 2];
///     char *operators[] = {
///         "containsText",
///         "notContains",
///         "beginsWith",
///         "endsWith",
///     };
///     uint8_t criteria = cond_format->criteria;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING)
///         LXW_PUSH_ATTRIBUTES_STR("type", "containsText");
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING)
///         LXW_PUSH_ATTRIBUTES_STR("type", "notContainsText");
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH)
///         LXW_PUSH_ATTRIBUTES_STR("type", "beginsWith");
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH)
///         LXW_PUSH_ATTRIBUTES_STR("type", "endsWith");
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     pos = criteria - LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING;
///     LXW_PUSH_ATTRIBUTES_STR("operator", operators[pos]);
///
///     LXW_PUSH_ATTRIBUTES_STR("text", cond_format->min_value_string);
///
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING) {
///         lxw_snprintf(formula, LXW_ATTR_32 * 2,
///                      "NOT(ISERROR(SEARCH(\"%s\",%s)))",
///                      cond_format->min_value_string, cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING) {
///         lxw_snprintf(formula, LXW_ATTR_32 * 2,
///                      "ISERROR(SEARCH(\"%s\",%s))",
///                      cond_format->min_value_string, cond_format->first_cell);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH) {
///         lxw_snprintf(formula, LXW_ATTR_32 * 2,
///                      "LEFT(%s,%d)=\"%s\"",
///                      cond_format->first_cell,
///                      (uint16_t) strlen(cond_format->min_value_string),
///                      cond_format->min_value_string);
///         _worksheet_write_formula_str(self, formula);
///     }
///     else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {
///         lxw_snprintf(formula, LXW_ATTR_32 * 2,
///                      "RIGHT(%s,%d)=\"%s\"",
///                      cond_format->first_cell,
///                      (uint16_t) strlen(cond_format->min_value_string),
///                      cond_format->min_value_string);
///         _worksheet_write_formula_str(self, formula);
///     }
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule_cell(lxw_worksheet *self,
///                               lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char *operators[] = {
///         "none",
///         "equal",
///         "notEqual",
///         "greaterThan",
///         "lessThan",
///         "greaterThanOrEqual",
///         "lessThanOrEqual",
///         "between",
///         "notBetween",
///     };
///
///     LXW_INIT_ATTRIBUTES();
///
///     LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
///
///     if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
///         LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);
///
///     LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
///
///     if (cond_format->stop_if_true)
///         LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);
///
///     LXW_PUSH_ATTRIBUTES_STR("operator", operators[cond_format->criteria]);
///
///     lxw_xml_start_tag(self->file, "cfRule", &attributes);
///
///     if (cond_format->min_value_string)
///         _worksheet_write_formula_str(self, cond_format->min_value_string);
///     else
///         _worksheet_write_formula_num(self, cond_format->min_value);
///
///     if (cond_format->has_max) {
///         if (cond_format->max_value_string)
///             _worksheet_write_formula_str(self, cond_format->max_value_string);
///         else
///             _worksheet_write_formula_num(self, cond_format->max_value);
///     }
///
///     lxw_xml_end_tag(self->file, "cfRule");
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_cf_rule(lxw_worksheet *self,
///                          lxw_cond_format_obj *cond_format)
/// {
///     if (cond_format->type == LXW_CONDITIONAL_TYPE_CELL) {
///
///         _worksheet_write_cf_rule_cell(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TEXT) {
///
///         _worksheet_write_cf_rule_text(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {
///
///         _worksheet_write_cf_rule_time_period(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_DUPLICATE
///              || cond_format->type == LXW_CONDITIONAL_TYPE_UNIQUE) {
///
///         _worksheet_write_cf_rule_duplicate(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_AVERAGE) {
///
///         _worksheet_write_cf_rule_average(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
///              || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {
///
///         _worksheet_write_cf_rule_top(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_BLANKS
///              || cond_format->type == LXW_CONDITIONAL_TYPE_NO_BLANKS
///              || cond_format->type == LXW_CONDITIONAL_TYPE_ERRORS
///              || cond_format->type == LXW_CONDITIONAL_TYPE_NO_ERRORS) {
///
///         _worksheet_write_cf_rule_blanks(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_FORMULA) {
///
///         _worksheet_write_cf_rule_formula(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_2_COLOR_SCALE
///              || cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
///
///         _worksheet_write_cf_rule_color_scale(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_DATA_BAR) {
///
///         _worksheet_write_cf_rule_data_bar(self, cond_format);
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_ICON_SETS) {
///
///         _worksheet_write_cf_rule_icons(self, cond_format);
///     }
///
/// }

/// STATIC void
/// _worksheet_write_conditional_formatting(lxw_worksheet *self,
///                                         lxw_cond_format_hash_element *element)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     lxw_cond_format_obj *cond_format;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("sqref", element->sqref);
///
///     lxw_xml_start_tag(self->file, "conditionalFormatting", &attributes);
///
///     STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
///         /* Write the cfRule element. */
///         _worksheet_write_cf_rule(self, cond_format);
///     }
///
///     lxw_xml_end_tag(self->file, "conditionalFormatting");
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_conditional_formats() const
{
  std::string xml_data;
  ///    lxw_cond_format_hash_element *element;
  ///     lxw_cond_format_hash_element *next_element;

  ///     for (element = RB_MIN(lxw_cond_format_hash,
  ///     self->conditional_formats);
  ///          element; element = next_element) {

  ///         _worksheet_write_conditional_formatting(self, element);

  ///         next_element =
  ///             RB_NEXT(lxw_cond_format_hash, self->conditional_formats,
  ///             element);
  ///     }

  return xml_data;
}

/// STATIC void
/// _worksheet_write_x14_color(lxw_worksheet *self, char *type, lxw_color_t color)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb[LXW_ATTR_32];
///
///     lxw_snprintf(rgb, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("rgb", rgb);
///     lxw_xml_empty_tag(self->file, type, &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_x14_cfvo(lxw_worksheet *self, uint8_t rule_type,
///                           double number, char *string)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char data[LXW_ATTR_32];
///     uint8_t has_value = LXW_FALSE;
///
///     LXW_INIT_ATTRIBUTES();
///
///     if (!string)
///         lxw_sprintf_dbl(data, number);
///
///     if (rule_type == LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "autoMin");
///         has_value = LXW_FALSE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "min");
///         has_value = LXW_FALSE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "num");
///         has_value = LXW_TRUE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "percent");
///         has_value = LXW_TRUE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
///         has_value = LXW_TRUE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "formula");
///         has_value = LXW_TRUE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "max");
///         has_value = LXW_FALSE;
///     }
///     else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX) {
///         LXW_PUSH_ATTRIBUTES_STR("type", "autoMax");
///         has_value = LXW_FALSE;
///     }
///
///     if (has_value) {
///         lxw_xml_start_tag(self->file, "x14:cfvo", &attributes);
///
///         if (string)
///             lxw_xml_data_element(self->file, "xm:f", string, NULL);
///         else
///             lxw_xml_data_element(self->file, "xm:f", data, NULL);
///
///         lxw_xml_end_tag(self->file, "x14:cfvo");
///     }
///     else {
///         lxw_xml_empty_tag(self->file, "x14:cfvo", &attributes);
///     }
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_x14_data_bar(lxw_worksheet *self,
///                               lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char min_length[] = "0";
///     char max_length[] = "100";
///     char border[] = "1";
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("minLength", min_length);
///     LXW_PUSH_ATTRIBUTES_STR("maxLength", max_length);
///
///     if (!cond_format->bar_no_border)
///         LXW_PUSH_ATTRIBUTES_STR("border", border);
///
///     if (cond_format->bar_solid)
///         LXW_PUSH_ATTRIBUTES_STR("gradient", "0");
///
///     if (cond_format->bar_direction ==
///         LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT)
///         LXW_PUSH_ATTRIBUTES_STR("direction", "rightToLeft");
///
///     if (cond_format->bar_direction ==
///         LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT)
///         LXW_PUSH_ATTRIBUTES_STR("direction", "leftToRight");
///
///     if (cond_format->bar_negative_color_same)
///         LXW_PUSH_ATTRIBUTES_STR("negativeBarColorSameAsPositive", "1");
///
///     if (!cond_format->bar_no_border
///         && !cond_format->bar_negative_border_color_same)
///         LXW_PUSH_ATTRIBUTES_STR("negativeBarBorderColorSameAsPositive", "0");
///
///     if (cond_format->bar_axis_position == LXW_CONDITIONAL_BAR_AXIS_MIDPOINT)
///         LXW_PUSH_ATTRIBUTES_STR("axisPosition", "middle");
///
///     if (cond_format->bar_axis_position == LXW_CONDITIONAL_BAR_AXIS_NONE)
///         LXW_PUSH_ATTRIBUTES_STR("axisPosition", "none");
///
///     lxw_xml_start_tag(self->file, "x14:dataBar", &attributes);
///
///     if (cond_format->auto_min)
///         cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN;
///
///     _worksheet_write_x14_cfvo(self, cond_format->min_rule_type,
///                               cond_format->min_value,
///                               cond_format->min_value_string);
///
///     if (cond_format->auto_max)
///         cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX;
///
///     _worksheet_write_x14_cfvo(self, cond_format->max_rule_type,
///                               cond_format->max_value,
///                               cond_format->max_value_string);
///
///     if (!cond_format->bar_no_border)
///         _worksheet_write_x14_color(self, "x14:borderColor",
///                                    cond_format->bar_border_color);
///
///     if (!cond_format->bar_negative_color_same)
///         _worksheet_write_x14_color(self, "x14:negativeFillColor",
///                                    cond_format->bar_negative_color);
///
///     if (!cond_format->bar_no_border
///         && !cond_format->bar_negative_border_color_same)
///         _worksheet_write_x14_color(self, "x14:negativeBorderColor",
///                                    cond_format->bar_negative_border_color);
///
///     if (cond_format->bar_axis_position != LXW_CONDITIONAL_BAR_AXIS_NONE)
///         _worksheet_write_x14_color(self, "x14:axisColor",
///                                    cond_format->bar_axis_color);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_x14_cf_rule(lxw_worksheet *self,
///                              lxw_cond_format_obj *cond_format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("type", "dataBar");
///     LXW_PUSH_ATTRIBUTES_STR("id", cond_format->guid);
///
///     lxw_xml_start_tag(self->file, "x14:cfRule", &attributes);
///
///     /* Write the x14:dataBar element. */
///     _worksheet_write_x14_data_bar(self, cond_format);
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_xm_sqref(lxw_worksheet *self,
///                           lxw_cond_format_obj *cond_format)
/// {
///     lxw_xml_data_element(self->file, "xm:sqref", cond_format->sqref, NULL);
/// }

/// STATIC void
/// _worksheet_write_conditional_formatting_2010(lxw_worksheet *self, lxw_cond_format_hash_element
///                                              *element)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     lxw_cond_format_obj *cond_format;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("xmlns:xm",
///                             "http://schemas.microsoft.com/office/excel/2006/main");
///
///     STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
///         if (!cond_format->data_bar_2010)
///             continue;
///
///         lxw_xml_start_tag(self->file, "x14:conditionalFormatting",
///                           &attributes);
///
///         _worksheet_write_x14_cf_rule(self, cond_format);
///
///         lxw_xml_end_tag(self->file, "x14:dataBar");
///         lxw_xml_end_tag(self->file, "x14:cfRule");
///         _worksheet_write_xm_sqref(self, cond_format);
///         lxw_xml_end_tag(self->file, "x14:conditionalFormatting");
///     }
///
///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _worksheet_write_ext_list_data_bars(lxw_worksheet *self)
/// {
///     lxw_cond_format_hash_element *element;
///     lxw_cond_format_hash_element *next_element;
///
///     _worksheet_write_ext(self, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
///     lxw_xml_start_tag(self->file, "x14:conditionalFormattings", NULL);
///
///     for (element = RB_MIN(lxw_cond_format_hash, self->conditional_formats);
///          element; element = next_element) {
///
///         _worksheet_write_conditional_formatting_2010(self, element);
///
///         next_element =
///             RB_NEXT(lxw_cond_format_hash, self->conditional_formats, element);
///     }
///
///     lxw_xml_end_tag(self->file, "x14:conditionalFormattings");
///     lxw_xml_end_tag(self->file, "ext");
/// }

std::string worksheet_t::write_ext_list() const
{
  std::string xml_data;
  ///     if (self->data_bar_2010_index == 0)
  ///         return;

  ///     lxw_xml_start_tag(self->file, "extLst", NULL);

  ///     _worksheet_write_ext_list_data_bars(self);

  ///     lxw_xml_end_tag(self->file, "extLst");

  return xml_data;
}

/// STATIC void
/// _worksheet_write_ignored_error(lxw_worksheet *self, char *ignore_error,
///                                char *range)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("sqref", range);
///     LXW_PUSH_ATTRIBUTES_STR(ignore_error, "1");
///
///     lxw_xml_empty_tag(self->file, "ignoredError", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }
///
/// lxw_error
/// _validate_conditional_icons(lxw_conditional_format *user)
/// {
///     if (user->icon_style > LXW_CONDITIONAL_ICONS_5_QUARTERS) {
///
///         LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                          "For type = LXW_CONDITIONAL_TYPE_ICON_SETS, "
///                          "invalid icon_style (%d).", user->icon_style);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     else {
///         return LXW_NO_ERROR;
///     }
/// }

/// lxw_error
/// _validate_conditional_data_bar(lxw_worksheet *self,
///                                lxw_cond_format_obj *cond_format,
///                                lxw_conditional_format *user_options)
/// {
///     uint8_t min_rule_type = user_options->min_rule_type;
///     uint8_t max_rule_type = user_options->max_rule_type;
///
///     if (user_options->data_bar_2010
///         || user_options->bar_solid
///         || user_options->bar_no_border
///         || user_options->bar_direction
///         || user_options->bar_axis_position
///         || user_options->bar_negative_color_same
///         || user_options->bar_negative_border_color_same
///         || user_options->bar_negative_color
///         || user_options->bar_border_color
///         || user_options->bar_negative_border_color
///         || user_options->bar_axis_color) {
///
///         cond_format->data_bar_2010 = LXW_TRUE;
///         self->excel_version = 2010;
///     }
///
///     if (min_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///         && min_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///         cond_format->min_rule_type = min_rule_type;
///         cond_format->min_value = user_options->min_value;
///         cond_format->min_value_string =
///             lxw_strdup_formula(user_options->min_value_string);
///     }
///     else {
///         cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_MINIMUM;
///         cond_format->min_value = 0;
///     }
///
///     if (max_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///         && max_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///         cond_format->max_rule_type = max_rule_type;
///         cond_format->max_value = user_options->max_value;
///         cond_format->max_value_string =
///             lxw_strdup_formula(user_options->max_value_string);
///     }
///     else {
///         cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_MAXIMUM;
///         cond_format->max_value = 0;
///     }
///
///     if (cond_format->data_bar_2010) {
///         if (min_rule_type == LXW_CONDITIONAL_RULE_TYPE_NONE)
///             cond_format->auto_min = LXW_TRUE;
///         if (max_rule_type == LXW_CONDITIONAL_RULE_TYPE_NONE)
///             cond_format->auto_max = LXW_TRUE;
///     }
///
///     cond_format->bar_only = user_options->bar_only;
///     cond_format->bar_solid = user_options->bar_solid;
///     cond_format->bar_no_border = user_options->bar_no_border;
///     cond_format->bar_direction = user_options->bar_direction;
///     cond_format->bar_axis_position = user_options->bar_axis_position;
///     cond_format->bar_negative_color_same =
///         user_options->bar_negative_color_same;
///     cond_format->bar_negative_border_color_same =
///         user_options->bar_negative_border_color_same;
///
///     if (user_options->bar_color != LXW_COLOR_UNSET)
///         cond_format->bar_color = user_options->bar_color;
///     else
///         cond_format->bar_color = 0x638EC6;
///
///     if (user_options->bar_negative_color != LXW_COLOR_UNSET)
///         cond_format->bar_negative_color = user_options->bar_negative_color;
///     else
///         cond_format->bar_negative_color = 0xFF0000;
///
///     if (user_options->bar_border_color != LXW_COLOR_UNSET)
///         cond_format->bar_border_color = user_options->bar_border_color;
///     else
///         cond_format->bar_border_color = cond_format->bar_color;
///
///     if (user_options->bar_negative_border_color != LXW_COLOR_UNSET)
///         cond_format->bar_negative_border_color =
///             user_options->bar_negative_border_color;
///     else
///         cond_format->bar_negative_border_color = 0xFF0000;
///
///     if (user_options->bar_axis_color != LXW_COLOR_UNSET)
///         cond_format->bar_axis_color = user_options->bar_axis_color;
///     else
///         cond_format->bar_axis_color = 0x000000;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// _validate_conditional_scale(lxw_cond_format_obj *cond_format,
///                             lxw_conditional_format *user_options)
/// {
///     uint8_t min_rule_type = user_options->min_rule_type;
///     uint8_t mid_rule_type = user_options->mid_rule_type;
///     uint8_t max_rule_type = user_options->max_rule_type;
///
///     if (min_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///         && min_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///         cond_format->min_rule_type = min_rule_type;
///         cond_format->min_value = user_options->min_value;
///         cond_format->min_value_string =
///             lxw_strdup_formula(user_options->min_value_string);
///     }
///     else {
///         cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_MINIMUM;
///         cond_format->min_value = 0;
///     }
///
///     if (max_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///         && max_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///         cond_format->max_rule_type = max_rule_type;
///         cond_format->max_value = user_options->max_value;
///         cond_format->max_value_string =
///             lxw_strdup_formula(user_options->max_value_string);
///     }
///     else {
///         cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_MAXIMUM;
///         cond_format->max_value = 0;
///     }
///
///     if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
///         if (mid_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
///             && mid_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
///             cond_format->mid_rule_type = mid_rule_type;
///             cond_format->mid_value = user_options->mid_value;
///             cond_format->mid_value_string =
///                 lxw_strdup_formula(user_options->mid_value_string);
///         }
///         else {
///             cond_format->mid_rule_type = LXW_CONDITIONAL_RULE_TYPE_PERCENTILE;
///             cond_format->mid_value = 50;
///         }
///     }
///
///     if (user_options->min_color != LXW_COLOR_UNSET)
///         cond_format->min_color = user_options->min_color;
///     else
///         cond_format->min_color = 0xFF7128;
///
///     if (user_options->max_color != LXW_COLOR_UNSET)
///         cond_format->max_color = user_options->max_color;
///     else
///         cond_format->max_color = 0xFFEF9C;
///
///     if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
///         if (user_options->min_color == LXW_COLOR_UNSET)
///             cond_format->min_color = 0xF8696B;
///
///         if (user_options->mid_color != LXW_COLOR_UNSET)
///             cond_format->mid_color = user_options->mid_color;
///         else
///             cond_format->mid_color = 0xFFEB84;
///
///         if (user_options->max_color == LXW_COLOR_UNSET)
///             cond_format->max_color = 0x63BE7B;
///     }
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// _validate_conditional_top(lxw_cond_format_obj *cond_format,
///                           lxw_conditional_format *user_options)
/// {
///     /* Restrict the range of rank values to Excel's allowed range. */
///     if (user_options->criteria ==
///         LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT) {
///         if (user_options->value < 0.0 || user_options->value > 100.0) {
///
///             LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                              "For type = LXW_CONDITIONAL_TYPE_TOP/BOTTOM, "
///                              "top/bottom percent (%g%%) must by in range 0-100",
///                              user_options->value);
///
///             return LXW_ERROR_PARAMETER_VALIDATION;
///         }
///     }
///     else {
///         if (user_options->value < 1.0 || user_options->value > 1000.0) {
///
///             LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                              "For type = LXW_CONDITIONAL_TYPE_TOP/BOTTOM, "
///                              "top/bottom items (%g) must by in range 1-1000",
///                              user_options->value);
///
///             return LXW_ERROR_PARAMETER_VALIDATION;
///         }
///     }
///
///     cond_format->min_value = (uint16_t) user_options->value;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// _validate_conditional_average(lxw_conditional_format *user)
/// {
///     if (user->criteria < LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE ||
///         user->criteria > LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW) {
///
///         LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                          "For type = LXW_CONDITIONAL_TYPE_AVERAGE, "
///                          "invalid criteria value (%d).", user->criteria);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     else {
///         return LXW_NO_ERROR;
///     }
/// }
///
/// lxw_error
/// _validate_conditional_time_period(lxw_conditional_format *user)
/// {
///     if (user->criteria < LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY ||
///         user->criteria > LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {
///
///         LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                          "For type = LXW_CONDITIONAL_TYPE_TIME_PERIOD, "
///                          "invalid criteria value (%d).", user->criteria);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     else {
///         return LXW_NO_ERROR;
///     }
/// }
///
/// lxw_error
/// _validate_conditional_text(lxw_cond_format_obj *cond_format,
///                            lxw_conditional_format *user_options)
/// {
///     if (!user_options->value_string) {
///
///         LXW_WARN_FORMAT("worksheet_conditional_format_cell()/_range(): "
///                         "For type = LXW_CONDITIONAL_TYPE_TEXT, "
///                         "value_string can not be NULL. "
///                         "Text must be specified.");
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (strlen(user_options->value_string) >= LXW_MAX_ATTRIBUTE_LENGTH) {
///
///         LXW_WARN_FORMAT2("worksheet_conditional_format_cell()/_range(): "
///                          "For type = LXW_CONDITIONAL_TYPE_TEXT, "
///                          "value_string length (%d) must be less than %d.",
///                          (uint16_t) strlen(user_options->value_string),
///                          LXW_MAX_ATTRIBUTE_LENGTH);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (user_options->criteria < LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING ||
///         user_options->criteria > LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {
///
///         LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                          "For type = LXW_CONDITIONAL_TYPE_TEXT, "
///                          "invalid criteria value (%d).",
///                          user_options->criteria);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     cond_format->min_value_string =
///         lxw_strdup_formula(user_options->value_string);
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// _validate_conditional_formula(lxw_cond_format_obj *cond_format,
///                               lxw_conditional_format *user_options)
/// {
///     if (!user_options->value_string) {
///
///         LXW_WARN_FORMAT("worksheet_conditional_format_cell()/_range(): "
///                         "For type = LXW_CONDITIONAL_TYPE_FORMULA, "
///                         "value_string can not be NULL. "
///                         "Formula must be specified.");
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     cond_format->min_value_string =
///         lxw_strdup_formula(user_options->value_string);
///
///     return LXW_NO_ERROR;
/// }
///
/// lxw_error
/// _validate_conditional_cell(lxw_cond_format_obj *cond_format,
///                            lxw_conditional_format *user_options)
/// {
///     cond_format->min_value = user_options->value;
///     cond_format->min_value_string =
///         lxw_strdup_formula(user_options->value_string);
///
///     if (cond_format->criteria == LXW_CONDITIONAL_CRITERIA_BETWEEN
///         || cond_format->criteria == LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN) {
///         cond_format->has_max = LXW_TRUE;
///         cond_format->min_value = user_options->min_value;
///         cond_format->max_value = user_options->max_value;
///         cond_format->min_value_string =
///             lxw_strdup_formula(user_options->min_value_string);
///         cond_format->max_value_string =
///             lxw_strdup_formula(user_options->max_value_string);
///     }
///
///     return LXW_NO_ERROR;
/// }
///
/// /* Check that the correct criteria and used with the correct conditional format. */
/// lxw_error
/// _validate_conditional_criteria(lxw_cond_format_obj *cond_format)
/// {
///     uint8_t criteria_mismatch = LXW_FALSE;
///
///     if (cond_format->type == LXW_CONDITIONAL_TYPE_CELL) {
///         switch (cond_format->criteria) {
///             case LXW_CONDITIONAL_CRITERIA_EQUAL_TO:
///             case LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO:
///             case LXW_CONDITIONAL_CRITERIA_GREATER_THAN:
///             case LXW_CONDITIONAL_CRITERIA_LESS_THAN:
///             case LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
///             case LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO:
///             case LXW_CONDITIONAL_CRITERIA_BETWEEN:
///             case LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN:
///                 criteria_mismatch = LXW_FALSE;
///                 break;
///             default:
///                 criteria_mismatch = LXW_TRUE;
///         }
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {
///         switch (cond_format->criteria) {
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH:
///             case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH:
///                 criteria_mismatch = LXW_FALSE;
///                 break;
///             default:
///                 criteria_mismatch = LXW_TRUE;
///         }
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TEXT) {
///         switch (cond_format->criteria) {
///             case LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING:
///             case LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING:
///             case LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH:
///             case LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH:
///                 criteria_mismatch = LXW_FALSE;
///                 break;
///             default:
///                 criteria_mismatch = LXW_TRUE;
///         }
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_AVERAGE) {
///         switch (cond_format->criteria) {
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE:
///             case LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW:
///                 criteria_mismatch = LXW_FALSE;
///                 break;
///             default:
///                 criteria_mismatch = LXW_TRUE;
///         }
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
///              || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {
///         switch (cond_format->criteria) {
///             case LXW_CONDITIONAL_CRITERIA_NONE:
///             case LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT:
///                 criteria_mismatch = LXW_FALSE;
///                 break;
///             default:
///                 criteria_mismatch = LXW_TRUE;
///         }
///     }
///     else {
///         /* Any other conditional type should have a zero criteria. */
///         cond_format->criteria = LXW_CONDITIONAL_CRITERIA_NONE;
///     }
///
///     if (criteria_mismatch) {
///         LXW_WARN_FORMAT2("worksheet_conditional_format_cell()/_range(): "
///                          "LXW_CONDITIONAL_CRITERIA_* = %d is not valid for "
///                          "LXW_CONDITIONAL_TYPE_* = %d", cond_format->criteria,
///                          cond_format->type);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     else {
///         return LXW_NO_ERROR;
///     }
/// }

std::string worksheet_t::write_ignored_errors() const
{
  std::string xml_data;

  ///   if (!self->has_ignore_errors)
  ///         return;

  ///     lxw_xml_start_tag(self->file, "ignoredErrors", NULL);

  ///     if (self->ignore_number_stored_as_text) {
  ///         _worksheet_write_ignored_error(self, "numberStoredAsText",
  ///                                        self->ignore_number_stored_as_text);
  ///     }

  ///     if (self->ignore_eval_error) {
  ///         _worksheet_write_ignored_error(self, "evalError",
  ///                                        self->ignore_eval_error);
  ///     }

  ///     if (self->ignore_formula_differs) {
  ///         _worksheet_write_ignored_error(self, "formula",
  ///                                       self->ignore_formula_differs);
  ///     }

  ///     if (self->ignore_formula_range) {
  ///         _worksheet_write_ignored_error(self, "formulaRange",
  ///                                        self->ignore_formula_range);
  ///     }

  ///     if (self->ignore_formula_unlocked) {
  ///         _worksheet_write_ignored_error(self, "unlockedFormula",
  ///                                        self->ignore_formula_unlocked);
  ///     }

  ///     if (self->ignore_empty_cell_reference) {
  ///         _worksheet_write_ignored_error(self, "emptyCellReference",
  ///                                        self->ignore_empty_cell_reference);
  ///     }

  ///     if (self->ignore_list_data_validation) {
  ///         _worksheet_write_ignored_error(self, "listDataValidation",
  ///                                        self->ignore_list_data_validation);
  ///     }

  ///     if (self->ignore_calculated_column) {
  ///         _worksheet_write_ignored_error(self, "calculatedColumn",
  ///                                        self->ignore_calculated_column);
  ///     }

  ///     if (self->ignore_two_digit_text_year) {
  ///         _worksheet_write_ignored_error(self, "twoDigitTextYear",
  ///                                        self->ignore_two_digit_text_year);
  ///     }

  ///     lxw_xml_end_tag(self->file, "ignoredErrors");

  return xml_data;
}

/// STATIC void
/// _worksheet_write_table_part(lxw_worksheet *self, uint16_t id)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char r_id[LXW_MAX_ATTRIBUTE_LENGTH];
///
///     lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);
///
///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);
///
///     lxw_xml_empty_tag(self->file, "tablePart", &attributes);
///
///     LXW_FREE_ATTRIBUTES();
/// }

std::string worksheet_t::write_table_parts() const
{
  std::string xml_data;
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     lxw_table_obj *table_obj;

  ///     if (!self->table_count)
  ///         return;

  ///     LXW_PUSH_ATTRIBUTES_INT("count", self->table_count);

  ///     lxw_xml_start_tag(self->file, "tableParts", &attributes);

  ///     STAILQ_FOREACH(table_obj, self->table_objs, list_pointers) {
  ///         self->rel_count++;

  /* Write the tablePart element. */
  ///         _worksheet_write_table_part(self, self->rel_count);
  ///     }

  ///     lxw_xml_end_tag(self->file, "tableParts");

  return xml_data;
}

/// void
/// lxw_worksheet_write_sheet_views(lxw_worksheet *self)
/// {
///     _worksheet_write_sheet_views(self);
/// }
///
/// void
/// lxw_worksheet_write_page_margins(lxw_worksheet *self)
/// {
///     _worksheet_write_page_margins(self);
/// }
///
/// void
/// lxw_worksheet_write_drawings(lxw_worksheet *self)
/// {
///     _worksheet_write_drawings(self);
/// }
///
/// void
/// lxw_worksheet_write_sheet_protection(lxw_worksheet *self,
///                                      lxw_protection_obj *protect)
/// {
///     _worksheet_write_sheet_protection(self, protect);
/// }
///
/// void
/// lxw_worksheet_write_sheet_pr(lxw_worksheet *self)
/// {
///     _worksheet_write_sheet_pr(self);
/// }
///
/// void
/// lxw_worksheet_write_page_setup(lxw_worksheet *self)
/// {
///     _worksheet_write_page_setup(self);
/// }
///
/// void
/// lxw_worksheet_write_header_footer(lxw_worksheet *self)
/// {
///     _worksheet_write_header_footer(self);
/// }

std::string worksheet_t::assemble_xml_file()
{
  std::string xml_data = xml_declaration();
  xml_data += write_worksheet();
  xml_data += write_sheet_pr();
  xml_data += write_dimension();
  xml_data += write_sheet_views();
  xml_data += write_sheet_format_pr();
  xml_data += write_cols();
  ///     if (!self->optimize)
  xml_data += write_sheet_data();
  ///     else
  ///         _worksheet_write_optimized_sheet_data(self);

  // Write the sheetProtection element.
  xml_data += write_sheet_protection();

  xml_data += write_auto_filter();
  xml_data += write_merge_cells();
  xml_data += write_conditional_formats();
  xml_data += write_data_validations();
  xml_data += write_hyperlinks();
  xml_data += write_print_options();
  xml_data += write_page_margins();
  xml_data += write_page_setup();
  xml_data += write_header_footer();
  xml_data += write_row_breaks();
  xml_data += write_col_breaks();
  xml_data += write_ignored_errors();
  xml_data += write_drawings();
  xml_data += write_legacy_drawing();
  xml_data += write_legacy_drawing_hf();
  xml_data += write_picture();
  xml_data += write_table_parts();
  xml_data += write_ext_list();
  xml_data += xml_end_tag("worksheet");

  return xml_data;
}

std::string worksheet_t::get_sheet_name() const
{
  return name_;
}

uint16_t worksheet_t::get_sheet_index() const
{
  return index_;
}

void worksheet_t::write_number(row_num_t row_num, col_num_t col_num, double number)
{
  write_number(row_num, col_num, number, nullptr);
}

void worksheet_t::write_number(row_num_t row_num, col_num_t col_num, double number, const format_t* format)
{
  check_dimensions(row_num, col_num, false, false);

  const cell_t cell = new_number_cell(row_num, col_num, number, format);

  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_string(row_num_t row_num, col_num_t col_num, const std::string& str)
{
  write_string(row_num, col_num, str, nullptr);
}

void worksheet_t::write_string(row_num_t row_num, col_num_t col_num, const std::string& str, const format_t* format)
{
  if(str.empty())
  {
    // Treat an empty string with formatting as a blank cell.
    // Empty strings without formats should be ignored.
    if(format)
    {
      write_blank(row_num, col_num, format);
    }

    return;
  }

  check_dimensions(row_num, col_num, false, false);

  if(str.size() > STR_MAX)
  {
    throw xwpp_out_of_range_t(std::format("String size {} too large (max {})", str.size(), STR_MAX));
  }

  ///     if (!self->optimize) {
  // Get the SST element and string id.
  const shared_strings_element_t sst_element = sst_->get_index(str, false);

  const cell_t cell = new_string_cell(row_num, col_num, sst_element.index_, sst_element.string_, format);
  ///     }
  ///     else {
  /* Look for and escape control chars in the string. */
  ///         if (lxw_has_control_characters(string)) {
  ///             string_copy = lxw_escape_control_characters(string);
  ///         }
  ///         else {
  ///             string_copy = lxw_strdup(string);
  ///         }
  ///         cell = _new_inline_string_cell(row_num, col_num, string_copy, format);
  ///     }

  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_formula_num(row_num_t row_num, col_num_t col_num, const std::string& formula, double result)
{
  write_formula_num(row_num, col_num, formula, nullptr, result);
}

void worksheet_t::write_formula_num(row_num_t row_num, col_num_t col_num, const std::string& formula,
                                    const format_t* format, double result)
{
  std::string formula_copy;

  if(formula.empty())
  {
    throw xwpp_exception_t("Formula is empty");
  }

  check_dimensions(row_num, col_num, false, false);

  // Strip leading "=" from formula.
  if(formula[0] == '=')
  {
    formula_copy = formula.substr(1);
  }
  else
  {
    formula_copy = formula;
  }

  cell_t cell = new_formula_cell(row_num, col_num, formula_copy, format, result);

  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_formula_str(row_num_t row_num, col_num_t col_num, const std::string& formula,
                                    const std::string& result)
{
  return write_formula_str(row_num, col_num, formula, nullptr, result);
}

void worksheet_t::write_formula_str(row_num_t row_num, col_num_t col_num, const std::string& formula,
                                    const format_t* format, const std::string& result)
{
  std::string formula_copy;

  if(formula.empty())
  {
    throw xwpp_exception_t("Formula is empty");
  }

  check_dimensions(row_num, col_num, false, false);

  // Strip leading "=" from formula.
  if(formula[0] == '=')
  {
    formula_copy = formula.substr(1);
  }
  else
  {
    formula_copy = formula;
  }

  cell_t cell = new_formula_cell(row_num, col_num, formula_copy, format, result);

  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_formula(row_num_t row_num, col_num_t col_num, const std::string& formula,
                                const format_t* format)
{
  write_formula_num(row_num, col_num, formula, format, 0);
}

void worksheet_t::write_formula(row_num_t row_num, col_num_t col_num, const std::string& formula)
{
  write_formula(row_num, col_num, formula, nullptr);
}

/// lxw_error
/// _store_array_formula(lxw_worksheet *self,
///                      row_num_t first_row,
///                      col_num_t first_col,
///                      row_num_t last_row,
///                      col_num_t last_col,
///                      const char *formula, lxw_format *format, double result,
///                      uint8_t is_dynamic)
/// {
///     cell_t *cell;
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     char *formula_copy;
///     char *range;
///     lxw_error err;
///
///     /* Swap last row/col with first row/col as necessary */
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     if (!formula)
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///
///     if (lxw_str_is_empty(formula))
///         return LXW_ERROR_PARAMETER_IS_EMPTY;
///
///     /* Check that row and col are valid and store max and min values. */
///     err = _check_dimensions(self, first_row, first_col, LXW_FALSE, LXW_FALSE);
///     if (err)
///         return err;
///
///     err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
///     if (err)
///         return err;
///
///     /* Define the array range. */
///     range = calloc(1, LXW_MAX_CELL_RANGE_LENGTH);
///     RETURN_ON_MEM_ERROR(range, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     if (first_row == last_row && first_col == last_col)
///         lxw_rowcol_to_cell(range, first_row, first_col);
///     else
///         lxw_rowcol_to_range(range, first_row, first_col, last_row, last_col);
///
///     /* Copy and trip leading "{=" from formula. */
///     if (formula[0] == '{')
///         if (strlen(formula) >= 2 && formula[1] == '=')
///             formula_copy = lxw_strdup(formula + 2);
///         else
///             formula_copy = lxw_strdup(formula + 1);
///     else
///         formula_copy = lxw_strdup_formula(formula);
///
///     /* Strip trailing "}" from formula. */
///     if (strlen(formula_copy) > 0
///         && formula_copy[strlen(formula_copy) - 1] == '}') {
///         formula_copy[strlen(formula_copy) - 1] = '\0';
///     }
///
///     /* Check for empty formula that started as {=}. */
///     if (lxw_str_is_empty(formula_copy)) {
///         free(formula_copy);
///         free(range);
///         return LXW_ERROR_PARAMETER_IS_EMPTY;
///     }
///
///     /* Create a new array formula cell object. */
///     cell = _new_array_formula_cell(first_row, first_col,
///                                    formula_copy, range, format, is_dynamic);
///
///     cell->formula_result = result;
///
///     _insert_cell(self, first_row, first_col, cell);
///
///     if (is_dynamic)
///         self->has_dynamic_functions = LXW_TRUE;
///
///     /* Pad out the rest of the area with formatted zeroes. */
///     if (!self->optimize) {
///         for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
///             for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
///                 if (tmp_row == first_row && tmp_col == first_col)
///                     continue;
///
///                 worksheet_write_number(self, tmp_row, tmp_col, 0, format);
///             }
///         }
///     }
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_write_array_formula_num(lxw_worksheet *self,
///                                   row_num_t first_row,
///                                   col_num_t first_col,
///                                   row_num_t last_row,
///                                   col_num_t last_col,
///                                   const char *formula,
///                                   lxw_format *format, double result)
/// {
///     return _store_array_formula(self, first_row, first_col,
///                                 last_row, last_col, formula, format, result,
///                                 LXW_FALSE);
/// }

/// lxw_error
/// worksheet_write_array_formula(lxw_worksheet *self,
///                               row_num_t first_row,
///                               col_num_t first_col,
///                               row_num_t last_row,
///                               col_num_t last_col,
///                               const char *formula, lxw_format *format)
/// {
///     return _store_array_formula(self, first_row, first_col,
///                                 last_row, last_col, formula, format, 0,
///                                 LXW_FALSE);
/// }

/// lxw_error
/// worksheet_write_dynamic_formula(lxw_worksheet *self,
///                                 row_num_t row,
///                                 col_num_t col,
///                                 const char *formula, lxw_format *format)
/// {
///     return _store_array_formula(self, row, col, row, col, formula, format, 0,
///                                 LXW_TRUE);
/// }

/// lxw_error
/// worksheet_write_dynamic_formula_num(lxw_worksheet *self,
///                                     row_num_t row,
///                                     col_num_t col,
///                                     const char *formula,
///                                     lxw_format *format, double result)
/// {
///     return _store_array_formula(self, row, col, row, col, formula, format,
///                                 result, LXW_TRUE);
/// }

/// lxw_error
/// worksheet_write_dynamic_array_formula_num(lxw_worksheet *self,
///                                           row_num_t first_row,
///                                           col_num_t first_col,
///                                           row_num_t last_row,
///                                           col_num_t last_col,
///                                           const char *formula,
///                                           lxw_format *format, double result)
/// {
///     return _store_array_formula(self, first_row, first_col,
///                                 last_row, last_col, formula, format, result,
///                                 LXW_TRUE);
/// }

/// lxw_error
/// worksheet_write_dynamic_array_formula(lxw_worksheet *self,
///                                       row_num_t first_row,
///                                       col_num_t first_col,
///                                       row_num_t last_row,
///                                       col_num_t last_col,
///                                       const char *formula, lxw_format *format)
/// {
///     return _store_array_formula(self, first_row, first_col,
///                                 last_row, last_col, formula, format, 0,
///                                 LXW_TRUE);
/// }

void worksheet_t::write_blank(row_num_t row_num, col_num_t col_num, const format_t* format)
{
  // Blank cells without formatting are ignored by Excel.
  if(!format)
  {
    return;
  }

  check_dimensions(row_num, col_num, false, false);

  cell_t cell = new_blank_cell(row_num, col_num, format);
  insert_cell(row_num, col_num, cell);
}

/// lxw_error
/// worksheet_write_boolean(lxw_worksheet *self,
///                         row_num_t row_num, col_num_t col_num,
///                         int value, lxw_format *format)
/// {
///     cell_t *cell;
///     lxw_error err;
///
///     err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
///     if (err)
///         return err;
///
///     cell = _new_boolean_cell(row_num, col_num, value, format);
///
///     _insert_cell(self, row_num, col_num, cell);
///
///     return LXW_NO_ERROR;
/// }

void worksheet_t::write_datetime(row_num_t row_num, col_num_t col_num,
                                 const std::chrono::system_clock::time_point& datetime)
{
  write_datetime(row_num, col_num, datetime, nullptr);
}

void worksheet_t::write_datetime(row_num_t row_num, col_num_t col_num,
                                 const std::chrono::system_clock::time_point& datetime, const format_t* format)
{
  check_dimensions(row_num, col_num, false, false);
  const double excel_date = datetime_to_excel_date_with_epoch(datetime, false /* TODOself->use_1904_epoch*/);
  const cell_t cell       = new_number_cell(row_num, col_num, excel_date, format);
  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_unixtime(row_num_t row_num, col_num_t col_num, int64_t unixtime)
{
  write_unixtime(row_num, col_num, unixtime, nullptr);
}

void worksheet_t::write_unixtime(row_num_t row_num, col_num_t col_num, int64_t unixtime, const format_t* format)
{
  check_dimensions(row_num, col_num, false, false);
  const double excel_date = unixtime_to_excel_date_with_epoch(unixtime, false /* TODO self->use_1904_epoch*/);
  const cell_t cell       = new_number_cell(row_num, col_num, excel_date, format);
  insert_cell(row_num, col_num, cell);
}

void worksheet_t::write_url(row_num_t row_num, col_num_t col_num, const std::string& url)
{
  write_url(row_num, col_num, url, nullptr, "", "");
}

void worksheet_t::write_url(row_num_t row_num, col_num_t col_num, const std::string& url, const format_t* format)
{
  write_url(row_num, col_num, url, format, "", "");
}

void worksheet_t::write_url(row_num_t row_num, col_num_t col_num, const std::string& url, const format_t* format,
                            const std::string& str, const std::string& tooltip)
{
  std::string string_copy;
  std::string url_copy;
  std::string url_external;
  std::string url_string;
  cell_types_t link_type = cell_types_t::HYPERLINK_URL;

  if(url.empty())
  {
    throw xwpp_exception_t("URL is empty");
  }

  // Check the Excel limit of URLS per worksheet.
  if(hlink_count_ > MAX_NUMBER_URLS)
  {
    throw xwpp_out_of_range_t(std::format("Max number of URL exceeded ({} / {})", hlink_count_, MAX_NUMBER_URLS));
  }

  check_dimensions(row_num, col_num, false, false);

  // Set the URI scheme from internal links.
  if(url.find("internal:") != std::string::npos)
  {
    link_type = cell_types_t::HYPERLINK_INTERNAL;
  }

  // Set the URI scheme from external links.
  if(url.find("external:") != std::string::npos)
  {
    link_type = cell_types_t::HYPERLINK_EXTERNAL;
  }

  if(!str.empty())
  {
    string_copy = str;
  }
  else
  {
    if(link_type == cell_types_t::HYPERLINK_URL)
    {
      // Strip the mailto header.
      if(url.find("mailto:") != std::string::npos)
      {
        string_copy = url.substr(sizeof("mailto"));
      }
      else
      {
        string_copy = url;
      }
    }
    else
    {
      string_copy = url.substr(sizeof("__ternal"));
    }
  }

  if(link_type == cell_types_t::HYPERLINK_URL)
  {
    url_copy = url;
  }
  else
  {
    url_copy = url.substr(sizeof("__ternal"));
  }

  if(link_type == cell_types_t::HYPERLINK_INTERNAL)
  {
    url_string = string_copy;
  }

  // Split url into the link and optional anchor/location.
  size_t found_string = url_copy.find("#");
  if(found_string != std::string::npos)
  {
    url_string = url_copy.substr(found_string + 1);
    url_copy   = url_copy.substr(0, found_string);
  }

  // Escape the URL.
  if(link_type == cell_types_t::HYPERLINK_URL || link_type == cell_types_t::HYPERLINK_EXTERNAL)
  {
    url_copy = escape_url_characters(url_copy, false);
  }

  if(link_type == cell_types_t::HYPERLINK_EXTERNAL)
  {
    // External Workbook links need to be modified into the right format.
    // The URL will look something like "c:\temp\file.xlsx#Sheet!A1".

    // For external links change the dir separator from Unix to DOS.
    for(auto& c: url_copy)
    {
      if(c == '/')
      {
        c = '\\';
      }
    }

    for(auto& c: string_copy)
    {
      if(c == '/')
      {
        c = '\\';
      }
    }

    // Look for Windows style "C:/" link or Windows share "\\" link.
    found_string = url_copy.find(':');
    if(found_string == std::string::npos)
    {
      found_string = url_copy.find("\\\\");
    }

    if(found_string != std::string::npos)
    {
      // Add the file:/// URI to the url if non-local.
      url_external = "file:///"s + url_copy;
    }

    // Convert a ./dir/file.xlsx link to dir/file.xlsx.
    found_string = url_copy.find(".\\");
    if(found_string == 0)
    {
      url_copy = url_copy.substr(2);
    }

    if(!url_external.empty())
    {
      url_copy = url_external;
    }
  }

  // Check if URL exceeds Excel's length limit.
  if(url_copy.size() > max_url_length_)
  {
    throw xwpp_out_of_range_t(std::format("URL too long ({} / {})", url_copy.size(), max_url_length_));
  }

  // Use the default URL format if none is specified.
  if(!storing_embedded_image_)
  {
    write_string(row_num, col_num, string_copy, format ? format : default_url_format_);
  }

  const cell_t link = new_hyperlink_cell(row_num, col_num, link_type, url_copy, url_string, tooltip);
  insert_hyperlink(row_num, col_num, link);
  hlink_count_++;
}

/*
 * Write a rich string to an Excel file.
 *
 * Rather than duplicate several of the styles.c font xml methods of styles.c
 * and write the data to a memory buffer this function creates a temporary
 * styles object and uses it to write the data to a file. It then reads that
 * data back into memory and closes the file.
 */
/// lxw_error
/// worksheet_write_rich_string(lxw_worksheet *self,
///                             row_num_t row_num,
///                             col_num_t col_num,
///                             lxw_rich_string_tuple *rich_strings[],
///                             lxw_format *format)
/// {
///     cell_t *cell;
///     int32_t string_id;
///     struct sst_element *sst_element;
///     lxw_error err;
///     uint8_t i;
///     long file_size;
///     char *rich_string = NULL;
///     const char *string_copy = NULL;
///     lxw_styles *styles = NULL;
///     lxw_format *default_format = NULL;
///     lxw_rich_string_tuple *rich_string_tuple = NULL;
///     FILE *tmpfile;
///
///     err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
///     if (err)
///         return err;
///
///     /* Iterate through rich string fragments to check for input errors. */
///     i = 0;
///     err = LXW_NO_ERROR;
///     while ((rich_string_tuple = rich_strings[i++]) != NULL) {
///
///         /* Check for NULL or empty strings. */
///         if (!rich_string_tuple->string || !*rich_string_tuple->string) {
///             err = LXW_ERROR_PARAMETER_VALIDATION;
///         }
///     }
///
///     /* If there are less than 2 fragments it isn't a rich string. */
///     if (i <= 2)
///         err = LXW_ERROR_PARAMETER_VALIDATION;
///
///     if (err)
///         return err;
///
///     /* Create a tmp file for the styles object. */
///     tmpfile = lxw_get_filehandle(&rich_string, NULL, self->tmpdir);
///     if (!tmpfile)
///         return LXW_ERROR_CREATING_TMPFILE;
///
///     /* Create a temp styles object for writing the font data. */
///     styles = lxw_styles_new();
///     GOTO_LABEL_ON_MEM_ERROR(styles, mem_error);
///     styles->file = tmpfile;
///
///     /* Create a default format for non-formatted text. */
///     default_format = lxw_format_new();
///     GOTO_LABEL_ON_MEM_ERROR(default_format, mem_error);
///
///     /* Iterate through the rich string fragments and write each one out. */
///     i = 0;
///     while ((rich_string_tuple = rich_strings[i++]) != NULL) {
///         lxw_xml_start_tag(tmpfile, "r", NULL);
///
///         if (rich_string_tuple->format) {
///             /* Write the user defined font format. */
///             lxw_styles_write_rich_font(styles, rich_string_tuple->format);
///         }
///         else {
///             /* Write a default font format. Except for the first fragment. */
///             if (i > 1)
///                 lxw_styles_write_rich_font(styles, default_format);
///         }
///
///         lxw_styles_write_string_fragment(styles, rich_string_tuple->string);
///         lxw_xml_end_tag(tmpfile, "r");
///     }
///
///     /* Free the temp objects. */
///     lxw_styles_free(styles);
///     lxw_format_free(default_format);
///
///     /* Flush the file. */
///     fflush(tmpfile);
///
///     if (!rich_string) {
///         /* Read the size to calculate the required memory. */
///         file_size = ftell(tmpfile);
///         /* Allocate a buffer for the rich string xml data. */
///         rich_string = calloc(file_size + 1, 1);
///         GOTO_LABEL_ON_MEM_ERROR(rich_string, mem_error);
///
///         /* Rewind the file and read the data into the memory buffer. */
///         rewind(tmpfile);
///         if (fread((void *) rich_string, file_size, 1, tmpfile) < 1) {
///             fclose(tmpfile);
///             free((void *) rich_string);
///             return LXW_ERROR_READING_TMPFILE;
///         }
///     }
///
///     /* Close the temp file. */
///     fclose(tmpfile);
///
///     if (lxw_utf8_strlen(rich_string) > LXW_STR_MAX) {
///         free((void *) rich_string);
///         return LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED;
///     }
///
///     if (!self->optimize) {
///         /* Get the SST element and string id. */
///         sst_element = lxw_get_sst_index(self->sst, rich_string, LXW_TRUE);
///         free((void *) rich_string);
///
///         if (!sst_element)
///             return LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND;
///
///         string_id = sst_element->index;
///         cell = _new_string_cell(row_num, col_num, string_id,
///                                 sst_element->string, format);
///     }
///     else {
///         /* Look for and escape control chars in the string. */
///         if (lxw_has_control_characters(rich_string)) {
///             string_copy = lxw_escape_control_characters(rich_string);
///             free((void *) rich_string);
///         }
///         else {
///             string_copy = rich_string;
///         }
///         cell = _new_inline_rich_string_cell(row_num, col_num, string_copy,
///                                             format);
///     }
///
///     _insert_cell(self, row_num, col_num, cell);
///
///     return LXW_NO_ERROR;
///
/// mem_error:
///     lxw_styles_free(styles);
///     lxw_format_free(default_format);
///     fclose(tmpfile);
///
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
/// }

void worksheet_t::write_comment(row_num_t row_num, col_num_t col_num, const std::string& text,
                                std::optional<comment_options_t> options)
{
  check_dimensions(row_num, col_num, false, false);

  if(text.empty())
  {
    throw xwpp_exception_t("Comment is empty");
  }

  if(text.size() > STR_MAX)
  {
    throw xwpp_out_of_range_t(std::format("Comment {} is too long ({} / {})", text, text.size(), STR_MAX));
  }

  vml_obj_t comment;
  comment.text_ = text;
  comment.row_  = row_num;
  comment.col_  = col_num;

  // Set user and default parameters for the comment.
  get_comment_params(comment, options);

  cell_t cell = new_comment_cell(row_num, col_num, comment);
  insert_comment(row_num, col_num, cell);

  has_vml_      = true;
  has_comments_ = true;

  // Insert a placeholder in the cell RB table in the same position so
  // that the worksheet row "spans" calculations are correct.
  insert_cell_placeholder(row_num, col_num);
}

void worksheet_t::write_comment(row_num_t row_num, col_num_t col_num, const std::string& text)
{
  write_comment(row_num, col_num, text, std::nullopt);
}

void worksheet_t::set_column(col_num_t first_col, col_num_t last_col, double width /* TODO,
                          lxw_format *format,
                          lxw_row_col_options *user_options*/)
{
  ///     lxw_col_options *copied_options;
  const bool ignore_row = true;
  const bool ignore_col = true;
  ///     uint8_t hidden = LXW_FALSE;
  ///     uint8_t level = 0;
  ///     uint8_t collapsed = LXW_FALSE;
  ///     col_num_t col;
  ///     lxw_error err;
  ///
  ///     if (user_options) {
  ///         hidden = user_options->hidden;
  ///         level = user_options->level;
  ///         collapsed = user_options->collapsed;
  ///     }
  ///
  // Ensure second col is larger than first.
  if(first_col > last_col)
  {
    std::swap(first_col, last_col);
  }

  /* Ensure that the cols are valid and store max and min values.
   * NOTE: The check shouldn't modify the row dimensions and should only
   *       modify the column dimensions in certain cases. */
  ///     if (TODO format != NULL || (width != LXW_DEF_COL_WIDTH && hidden))
  ///         ignore_col = LXW_FALSE;
  ///
  check_dimensions(0, first_col, ignore_row, ignore_col);
  check_dimensions(0, last_col, ignore_row, ignore_col);

  if(first_col >= col_options_.size())
  {
    col_options_.resize(first_col + 1);
  }
  ///
  ///     /* Store the column options. */
  ///     copied_options = calloc(1, sizeof(lxw_col_options));
  ///     RETURN_ON_MEM_ERROR(copied_options, LXW_ERROR_MEMORY_MALLOC_FAILED);
  ///
  ///     /* Ensure the level is <= 7). */
  ///     if (level > 7)
  ///         level = 7;
  ///
  ///     if (level > self->outline_col_level)
  ///         self->outline_col_level = level;
  ///
  // Set the column properties.
  col_options_[first_col].firstcol_ = first_col;
  col_options_[first_col].lastcol_  = last_col;
  col_options_[first_col].width_    = width;
  ///     copied_options->format = format;
  ///     copied_options->hidden = hidden;
  ///     copied_options->level = level;
  ///     copied_options->collapsed = collapsed;
  ///
  ///     /* Store the column formats for use when writing cell data. */
  ///     for (col = firstcol; col <= lastcol; col++) {
  ///         self->col_formats[col] = format;
  ///     }
  ///
  // Store the column change to allow optimizations.
  col_size_changed_                 = true;
}

/// lxw_error
/// worksheet_set_column(lxw_worksheet *self,
///                      col_num_t firstcol,
///                      col_num_t lastcol, double width, lxw_format *format)
/// {
///     return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
///                                     NULL);
/// }

/// lxw_error
/// worksheet_set_column_pixels(lxw_worksheet *self,
///                             col_num_t firstcol,
///                             col_num_t lastcol,
///                             uint32_t pixels, lxw_format *format)
/// {
///     double width = _pixels_to_width(pixels);
///
///     return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
///                                     NULL);
/// }

/// lxw_error
/// worksheet_set_column_pixels_opt(lxw_worksheet *self,
///                                 col_num_t firstcol,
///                                 col_num_t lastcol,
///                                 uint32_t pixels,
///                                 lxw_format *format,
///                                 lxw_row_col_options *user_options)
/// {
///     double width = _pixels_to_width(pixels);
///
///     return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
///                                     user_options);
/// }

void worksheet_t::set_row(row_num_t row_num,
                          double height /*TODO, lxw_format *format, lxw_row_col_options *user_options*/)
{
  const col_num_t min_col = (dim_colmin_ != COL_MAX ? dim_colmin_ : 0);
  ///     uint8_t hidden = LXW_FALSE;
  ///     uint8_t level = 0;
  ///     uint8_t collapsed = LXW_FALSE;
  ///     row_t *row;
  ///     lxw_error err;
  ///
  ///     if (user_options) {
  ///         hidden = user_options->hidden;
  ///         level = user_options->level;
  ///         collapsed = user_options->collapsed;
  ///     }

  check_dimensions(row_num, min_col, false, false);
  ///
  ///     /* If the height is 0 the row is hidden and the height is the default. */
  ///     if (height == 0) {
  ///         hidden = LXW_TRUE;
  ///         height = self->default_row_height;
  ///     }
  ///
  ///     /* Ensure the level is <= 7). */
  ///     if (level > 7)
  ///         level = 7;
  ///
  ///     if (level > self->outline_row_level)
  ///         self->outline_row_level = level;
  ///
  // Store the row properties.
  row_t& row  = get_row(row_num);
  row.height_ = height;
  ///     row->format = format;
  ///     row->hidden = hidden;
  ///     row->level = level;
  ///     row->collapsed = collapsed;
  ///     row->row_changed = LXW_TRUE;
  ///
  if(row.height_ != default_row_height_)
  {
    row.height_changed_ = true;
  }
}

/// lxw_error
/// worksheet_set_row(lxw_worksheet *self,
///                   row_num_t row_num, double height, lxw_format *format)
/// {
///     return worksheet_set_row_opt(self, row_num, height, format, NULL);
/// }

/// lxw_error
/// worksheet_set_row_pixels(lxw_worksheet *self,
///                          row_num_t row_num, uint32_t pixels,
///                          lxw_format *format)
/// {
///     double height = _pixels_to_height(pixels);
///
///     return worksheet_set_row_opt(self, row_num, height, format, NULL);
/// }

/// lxw_error
/// worksheet_set_row_pixels_opt(lxw_worksheet *self,
///                              row_num_t row_num,
///                              uint32_t pixels,
///                              lxw_format *format,
///                              lxw_row_col_options *user_options)
/// {
///     double height = _pixels_to_height(pixels);
///
///     return worksheet_set_row_opt(self, row_num, height, format, user_options);
/// }

void worksheet_t::merge_range(row_num_t first_row, col_num_t first_col, row_num_t last_row, col_num_t last_col,
                              const std::string& str, const format_t* format)
{
  ///     lxw_merged_range *merged_range;
  ///     row_num_t tmp_row;
  ///     col_num_t tmp_col;
  ///     lxw_error err;
  ///
  // Excel doesn't allow a single cell to be merged
  if(first_row == last_row && first_col == last_col)
  {
    throw xwpp_exception_t("Cannot merge one single cell");
  }

  // Swap last row/col with first row/col as necessary
  if(first_row > last_row)
  {
    std::swap(first_row, last_row);
  }
  if(first_col > last_col)
  {
    std::swap(first_col, last_col);
  }

  // Check that column number is valid and store the max value
  check_dimensions(last_row, last_col, false, false);

  // Store the merge range.
  merged_range_t merged_range{
      .first_row_ = first_row,
      .last_row_  = last_row,
      .first_col_ = first_col,
      .last_col_  = last_col,
  };

  merged_ranges_.push_back(merged_range);

  // Write the first cell
  write_string(first_row, first_col, str, format);

  // Pad out the rest of the area with formatted blank cells.
  for(row_num_t tmp_row = first_row; tmp_row <= last_row; tmp_row++)
  {
    for(col_num_t tmp_col = first_col; tmp_col <= last_col; tmp_col++)
    {
      if(tmp_row != first_row || tmp_col != first_col)
      {
        write_blank(tmp_row, tmp_col, format);
      }
    }
  }
}

/// lxw_error
/// worksheet_autofilter(lxw_worksheet *self, row_num_t first_row,
///                      col_num_t first_col, row_num_t last_row,
///                      col_num_t last_col)
/// {
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     lxw_error err;
///     lxw_filter_rule_obj **filter_rules;
///     col_num_t num_filter_rules;
///
///     /* Swap last row/col with first row/col as necessary */
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     /* Check that column number is valid and store the max value */
///     err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
///     if (err)
///         return err;
///
///     /* Create a array to hold filter rules. */
///     self->autofilter.in_use = LXW_FALSE;
///     self->autofilter.has_rules = LXW_FALSE;
///     _free_filter_rules(self);
///     num_filter_rules = last_col - first_col + 1;
///     filter_rules = calloc(num_filter_rules, sizeof(lxw_filter_rule_obj *));
///     RETURN_ON_MEM_ERROR(filter_rules, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     self->autofilter.in_use = LXW_TRUE;
///     self->autofilter.first_row = first_row;
///     self->autofilter.first_col = first_col;
///     self->autofilter.last_row = last_row;
///     self->autofilter.last_col = last_col;
///
///     self->filter_rules = filter_rules;
///     self->num_filter_rules = num_filter_rules;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_filter_column(lxw_worksheet *self, col_num_t col,
///                         lxw_filter_rule *rule)
/// {
///     lxw_filter_rule_obj *rule_obj;
///     uint16_t rule_index;
///
///     if (!rule) {
///         LXW_WARN("worksheet_filter_column(): rule parameter cannot be NULL");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (self->autofilter.in_use == LXW_FALSE) {
///         LXW_WARN("worksheet_filter_column(): "
///                  "Worksheet autofilter range hasn't been defined. "
///                  "Use worksheet_autofilter() first.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
///         LXW_WARN_FORMAT3("worksheet_filter_column(): "
///                          "Column '%d' is outside autofilter range "
///                          "'%d <= col <= %d'.", col,
///                          self->autofilter.first_col,
///                          self->autofilter.last_col);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Free any previous rule in the column slot. */
///     rule_index = col - self->autofilter.first_col;
///     _free_filter_rule(self->filter_rules[rule_index]);
///
///     /* Create a new rule and copy user input. */
///     rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
///     RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     rule_obj->col_num = rule_index;
///     rule_obj->type = LXW_FILTER_TYPE_SINGLE;
///     rule_obj->criteria1 = rule->criteria;
///     rule_obj->value1 = rule->value;
///
///     if (rule_obj->criteria1 != LXW_FILTER_CRITERIA_NON_BLANKS) {
///         rule_obj->value1_string = lxw_strdup(rule->value_string);
///     }
///     else {
///         rule_obj->criteria1 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
///         rule_obj->value1_string = lxw_strdup(" ");
///     }
///
///     if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
///         rule_obj->has_blanks = LXW_TRUE;
///
///     _set_custom_filter(rule_obj);
///
///     self->filter_rules[rule_index] = rule_obj;
///     self->filter_on = LXW_TRUE;
///     self->autofilter.has_rules = LXW_TRUE;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_filter_column2(lxw_worksheet *self, col_num_t col,
///                          lxw_filter_rule *rule1, lxw_filter_rule *rule2,
///                          uint8_t and_or)
/// {
///     lxw_filter_rule_obj *rule_obj;
///     uint16_t rule_index;
///
///     if (!rule1 || !rule2) {
///         LXW_WARN("worksheet_filter_column2(): rule parameter cannot be NULL");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (self->autofilter.in_use == LXW_FALSE) {
///         LXW_WARN("worksheet_filter_column2(): "
///                  "Worksheet autofilter range hasn't been defined. "
///                  "Use worksheet_autofilter() first.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
///         LXW_WARN_FORMAT3("worksheet_filter_column2(): "
///                          "Column '%d' is outside autofilter range "
///                          "'%d <= col <= %d'.", col,
///                          self->autofilter.first_col,
///                          self->autofilter.last_col);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Free any previous rule in the column slot. */
///     rule_index = col - self->autofilter.first_col;
///     _free_filter_rule(self->filter_rules[rule_index]);
///
///     /* Create a new rule and copy user input. */
///     rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
///     RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     if (and_or == LXW_FILTER_AND)
///         rule_obj->type = LXW_FILTER_TYPE_AND;
///     else
///         rule_obj->type = LXW_FILTER_TYPE_OR;
///
///     rule_obj->col_num = rule_index;
///
///     rule_obj->criteria1 = rule1->criteria;
///     rule_obj->value1 = rule1->value;
///
///     rule_obj->criteria2 = rule2->criteria;
///     rule_obj->value2 = rule2->value;
///
///     if (rule_obj->criteria1 != LXW_FILTER_CRITERIA_NON_BLANKS) {
///         rule_obj->value1_string = lxw_strdup(rule1->value_string);
///     }
///     else {
///         rule_obj->criteria1 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
///         rule_obj->value1_string = lxw_strdup(" ");
///     }
///
///     if (rule_obj->criteria2 != LXW_FILTER_CRITERIA_NON_BLANKS) {
///         rule_obj->value2_string = lxw_strdup(rule2->value_string);
///     }
///     else {
///         rule_obj->criteria2 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
///         rule_obj->value2_string = lxw_strdup(" ");
///     }
///
///     if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
///         rule_obj->has_blanks = LXW_TRUE;
///
///     if (rule_obj->criteria2 == LXW_FILTER_CRITERIA_BLANKS)
///         rule_obj->has_blanks = LXW_TRUE;
///
///     _set_custom_filter(rule_obj);
///
///     self->filter_rules[rule_index] = rule_obj;
///     self->filter_on = LXW_TRUE;
///     self->autofilter.has_rules = LXW_TRUE;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_filter_list(lxw_worksheet *self, col_num_t col, const char **list)
/// {
///     lxw_filter_rule_obj *rule_obj;
///     uint16_t rule_index;
///     uint8_t has_blanks = LXW_FALSE;
///     uint16_t num_filters = 0;
///     uint16_t input_list_index;
///     uint16_t rule_obj_list_index;
///     const char *str;
///     char **tmp_list;
///
///     if (!list) {
///         LXW_WARN("worksheet_filter_list(): list parameter cannot be NULL");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (self->autofilter.in_use == LXW_FALSE) {
///         LXW_WARN("worksheet_filter_list(): "
///                  "Worksheet autofilter range hasn't been defined. "
///                  "Use worksheet_autofilter() first.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
///         LXW_WARN_FORMAT3("worksheet_filter_list(): "
///                          "Column '%d' is outside autofilter range "
///                          "'%d <= col <= %d'.", col,
///                          self->autofilter.first_col,
///                          self->autofilter.last_col);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Count the number of non "Blanks" strings in the input list. */
///     input_list_index = 0;
///     while ((str = list[input_list_index]) != NULL) {
///         if (strncmp(str, "Blanks", 6) == 0)
///             has_blanks = LXW_TRUE;
///         else
///             num_filters++;
///
///         input_list_index++;
///     }
///
///     /* There should be at least one filter string. */
///     if (num_filters == 0) {
///         LXW_WARN("worksheet_filter_list(): "
///                  "list must have at least 1 non-blanks item.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Free any previous rule in the column slot. */
///     rule_index = col - self->autofilter.first_col;
///     _free_filter_rule(self->filter_rules[rule_index]);
///
///     /* Create a new rule and copy user input. */
///     rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
///     RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     tmp_list = calloc(num_filters + 1, sizeof(char *));
///     GOTO_LABEL_ON_MEM_ERROR(tmp_list, mem_error);
///
///     /* Copy input list (without any "Blanks" command) to an internal list. */
///     input_list_index = 0;
///     rule_obj_list_index = 0;
///     while ((str = list[input_list_index]) != NULL) {
///         if (strncmp(str, "Blanks", 6) != 0) {
///             tmp_list[rule_obj_list_index] = lxw_strdup(str);
///             rule_obj_list_index++;
///         }
///
///         input_list_index++;
///     }
///
///     rule_obj->list = tmp_list;
///     rule_obj->num_list_filters = num_filters;
///     rule_obj->is_custom = LXW_FALSE;
///     rule_obj->col_num = rule_index;
///     rule_obj->type = LXW_FILTER_TYPE_STRING_LIST;
///     rule_obj->has_blanks = has_blanks;
///
///     self->filter_rules[rule_index] = rule_obj;
///     self->filter_on = LXW_TRUE;
///     self->autofilter.has_rules = LXW_TRUE;
///
///     return LXW_NO_ERROR;
///
/// mem_error:
///     free(rule_obj);
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
///
/// }

/// lxw_error
/// worksheet_add_table(lxw_worksheet *self, row_num_t first_row,
///                     col_num_t first_col, row_num_t last_row,
///                     col_num_t last_col, lxw_table_options *user_options)
/// {
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     col_num_t num_cols;
///     lxw_error err;
///     lxw_table_obj *table_obj;
///     lxw_table_column **columns;
///
///     if (self->optimize) {
///         LXW_WARN_FORMAT("worksheet_add_table(): "
///                         "worksheet tables aren't supported in "
///                         "'constant_memory' mode");
///         return LXW_ERROR_FEATURE_NOT_SUPPORTED;
///     }
///
///     /* Swap last row/col with first row/col as necessary */
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     /* Check that column number is valid and store the max value */
///     err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
///     if (err)
///         return err;
///
///     num_cols = last_col - first_col + 1;
///
///     /* Check that there are sufficient data rows. */
///     err = _check_table_rows(first_row, last_row, user_options);
///     if (err)
///         return err;
///
///     /* Check that the the table name is valid. */
///     err = _check_table_name(user_options);
///     if (err)
///         return err;
///
///     /* Create a table object to copy from the user options. */
///     table_obj = calloc(1, sizeof(lxw_table_obj));
///     RETURN_ON_MEM_ERROR(table_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     columns = calloc(num_cols, sizeof(lxw_table_column *));
///     GOTO_LABEL_ON_MEM_ERROR(columns, error);
///
///     table_obj->columns = columns;
///     table_obj->num_cols = num_cols;
///     table_obj->first_row = first_row;
///     table_obj->first_col = first_col;
///     table_obj->last_row = last_row;
///     table_obj->last_col = last_col;
///
///     err = _set_default_table_columns(table_obj);
///     if (err)
///         goto error;
///
///     /* Create the table range. */
///     lxw_rowcol_to_range(table_obj->sqref,
///                         first_row, first_col, last_row, last_col);
///     lxw_rowcol_to_range(table_obj->filter_sqref,
///                         first_row, first_col, last_row, last_col);
///
///     /* Validate and copy user options to an internal object. */
///     if (user_options) {
///
///         _check_and_copy_table_style(table_obj, user_options);
///
///         table_obj->total_row = user_options->total_row;
///         table_obj->last_column = user_options->last_column;
///         table_obj->first_column = user_options->first_column;
///         table_obj->no_autofilter = user_options->no_autofilter;
///         table_obj->no_header_row = user_options->no_header_row;
///         table_obj->no_banded_rows = user_options->no_banded_rows;
///         table_obj->banded_columns = user_options->banded_columns;
///
///         if (user_options->no_header_row)
///             table_obj->no_autofilter = LXW_TRUE;
///
///         if (user_options->columns) {
///             err = _set_custom_table_columns(table_obj, user_options);
///             if (err)
///                 goto error;
///         }
///
///         if (user_options->total_row) {
///             lxw_rowcol_to_range(table_obj->filter_sqref,
///                                 first_row, first_col, last_row - 1, last_col);
///         }
///
///         if (user_options->name) {
///             table_obj->name = lxw_strdup(user_options->name);
///             if (!table_obj->name) {
///                 err = LXW_ERROR_MEMORY_MALLOC_FAILED;
///                 goto error;
///             }
///         }
///     }
///
///     _write_table_column_data(self, table_obj);
///
///     STAILQ_INSERT_TAIL(self->table_objs, table_obj, list_pointers);
///     self->table_count++;
///
///     return LXW_NO_ERROR;
///
/// error:
///     _free_worksheet_table(table_obj);
///     return err;
///
/// }

void worksheet_t::select()
{
  selected_ = true;

  // Selected worksheet can't be hidden.
  hidden_ = false;
}

/// void
/// worksheet_activate(lxw_worksheet *self)
/// {
///     self->selected = LXW_TRUE;
///     self->active = LXW_TRUE;
///
///     /* Active worksheet can't be hidden. */
///     self->hidden = LXW_FALSE;
///
///     *self->active_sheet = self->index;
/// }

/// void
/// worksheet_set_first_sheet(lxw_worksheet *self)
/// {
///     /* Active worksheet can't be hidden. */
///     self->hidden = LXW_FALSE;
///
///     *self->first_sheet = self->index;
/// }

/// void
/// worksheet_hide(lxw_worksheet *self)
/// {
///     self->hidden = LXW_TRUE;
///
///     /* A hidden worksheet shouldn't be active or selected. */
///     self->selected = LXW_FALSE;
///
///     /* If this is active_sheet or first_sheet reset the workbook value. */
///     if (*self->first_sheet == self->index)
///         *self->first_sheet = 0;
///
///     if (*self->active_sheet == self->index)
///         *self->active_sheet = 0;
/// }

/// lxw_error
/// worksheet_set_selection(lxw_worksheet *self,
///                         row_num_t first_row, col_num_t first_col,
///                         row_num_t last_row, col_num_t last_col)
/// {
///     lxw_selection *selection;
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     lxw_error err;
///     char active_cell[LXW_MAX_CELL_RANGE_LENGTH];
///     char sqref[LXW_MAX_CELL_RANGE_LENGTH];
///
///     /* Only allow selection to be set once to avoid freeing/re-creating it. */
///     if (!STAILQ_EMPTY(self->selections))
///         return LXW_ERROR_PARAMETER_VALIDATION;
///
///     /* Excel doesn't set a selection for cell A1 since it is the default. */
///     if (first_row == 0 && first_col == 0 && last_row == 0 && last_col == 0)
///         return LXW_NO_ERROR;
///
///     selection = calloc(1, sizeof(lxw_selection));
///     RETURN_ON_MEM_ERROR(selection, LXW_ERROR_MEMORY_MALLOC_FAILED);
///
///     /* Check that row and col are valid without storing. */
///     err = _check_dimensions(self, first_row, first_col, LXW_TRUE, LXW_TRUE);
///     if (err) {
///         free(selection);
///         return err;
///     }
///
///     err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
///     if (err) {
///         free(selection);
///         return err;
///     }
///
///     /* Set the cell range selection. Do this before swapping max/min to  */
///     /* allow the selection direction to be reversed. */
///     lxw_rowcol_to_cell(active_cell, first_row, first_col);
///
///     /* Swap last row/col for first row/col if necessary. */
///     if (first_row > last_row) {
///         tmp_row = first_row;
///         first_row = last_row;
///         last_row = tmp_row;
///     }
///
///     if (first_col > last_col) {
///         tmp_col = first_col;
///         first_col = last_col;
///         last_col = tmp_col;
///     }
///
///     /* If the first and last cell are the same write a single cell. */
///     if ((first_row == last_row) && (first_col == last_col))
///         lxw_rowcol_to_cell(sqref, first_row, first_col);
///     else
///         lxw_rowcol_to_range(sqref, first_row, first_col, last_row, last_col);
///
///     lxw_strcpy(selection->pane, "");
///     lxw_strcpy(selection->active_cell, active_cell);
///     lxw_strcpy(selection->sqref, sqref);
///
///     STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
///
///     return LXW_NO_ERROR;
/// }

/// void
/// worksheet_set_top_left_cell(lxw_worksheet *self, row_num_t row, col_num_t col)
/// {
///     if (row == 0 && col == 0)
///         return;
///
///     lxw_rowcol_to_cell(self->top_left_cell, row, col);
/// }

/// void
/// worksheet_freeze_panes_opt(lxw_worksheet *self,
///                            row_num_t first_row, col_num_t first_col,
///                            row_num_t top_row, col_num_t left_col,
///                            uint8_t type)
/// {
///     self->panes.first_row = first_row;
///     self->panes.first_col = first_col;
///     self->panes.top_row = top_row;
///     self->panes.left_col = left_col;
///     self->panes.x_split = 0.0;
///     self->panes.y_split = 0.0;
///
///     if (type)
///         self->panes.type = FREEZE_SPLIT_PANES;
///     else
///         self->panes.type = FREEZE_PANES;
/// }

/// void
/// worksheet_freeze_panes(lxw_worksheet *self,
///                        row_num_t first_row, col_num_t first_col)
/// {
///     worksheet_freeze_panes_opt(self, first_row, first_col,
///                                first_row, first_col, 0);
/// }

/// void
/// worksheet_split_panes_opt(lxw_worksheet *self,
///                           double y_split, double x_split,
///                           row_num_t top_row, col_num_t left_col)
/// {
///     self->panes.first_row = 0;
///     self->panes.first_col = 0;
///     self->panes.top_row = top_row;
///     self->panes.left_col = left_col;
///     self->panes.x_split = x_split;
///     self->panes.y_split = y_split;
///     self->panes.type = SPLIT_PANES;
/// }

/// void
/// worksheet_split_panes(lxw_worksheet *self, double y_split, double x_split)
/// {
///     worksheet_split_panes_opt(self, y_split, x_split, 0, 0);
/// }

/// void
/// worksheet_set_portrait(lxw_worksheet *self)
/// {
///     self->orientation = LXW_PORTRAIT;
///     self->page_setup_changed = LXW_TRUE;
/// }

/// void
/// worksheet_set_landscape(lxw_worksheet *self)
/// {
///     self->orientation = LXW_LANDSCAPE;
///     self->page_setup_changed = LXW_TRUE;
/// }

/// void
/// worksheet_set_page_view(lxw_worksheet *self)
/// {
///     self->page_view = LXW_TRUE;
/// }

/// void
/// worksheet_set_paper(lxw_worksheet *self, uint8_t paper_size)
/// {
///     if (paper_size > 118) {
///         LXW_WARN_FORMAT1("worksheet_set_paper(): invalid paper size: %d. "
///                          "Valid range is 0-118", paper_size);
///         return;
///     }
///
///     self->paper_size = paper_size;
///     self->page_setup_changed = LXW_TRUE;
/// }

/// void
/// worksheet_print_across(lxw_worksheet *self)
/// {
///     self->page_order = LXW_PRINT_ACROSS;
///     self->page_setup_changed = LXW_TRUE;
/// }

void worksheet_t::set_margins(double left, double right, double top, double bottom)
{
  if(left >= 0)
  {
    margin_left_ = left;
  }

  if(right >= 0)
  {
    margin_right_ = right;
  }

  if(top >= 0)
  {
    margin_top_ = top;
  }

  if(bottom >= 0)
  {
    margin_bottom_ = bottom;
  }
}

void worksheet_t::set_header(const std::string& str, const std::optional<header_footer_options_t>& options)
{
  ///     lxw_error err;
  ///     char *tmp_header;
  ///     char *found_string;
  ///     char *offset_string;

  if(str.empty())
  {
    throw xwpp_exception_t("Header is empty");
  }

  if(str.size() > HEADER_FOOTER_MAX)
  {
    throw xwpp_out_of_range_t(std::format("Header {} is empty too long ({}/{})", str, str.size(), HEADER_FOOTER_MAX));
  }

  std::string tmp_header = str;

  // Count &G placeholders and ensure there are sufficient images.
  uint8_t placeholder_count = 0;
  for(size_t i = 0; i < tmp_header.size() - 1; ++i)
  {
    if(tmp_header[i] == '&' && tmp_header[i + 1] == 'G')
    {
      placeholder_count++;
    }
  }

  if(placeholder_count > 0 && !options)
  {
    throw xwpp_exception_t("&G placeholders present but no image supplied");
  }

  if(options)
  {
    uint8_t image_count = 0;

    /* Ensure there are enough images to match the placeholders. There is
     * a potential bug where there are sufficient images but in the wrong
     * positions but we don't currently try to deal with that.*/
    if(!options->image_left_.empty())
    {
      image_count++;
    }
    if(!options->image_center_.empty())
    {
      image_count++;
    }
    if(!options->image_right_.empty())
    {
      image_count++;
    }

    if(placeholder_count != image_count)
    {
      throw xwpp_exception_t("Number of &G placeholders does not match number of supplied images");
    }

    if(options->margin_ > 0.0)
    {
      margin_header_ = options->margin_;
    }

    set_header_footer_image(options->image_left_, image_position_t::HEADER_LEFT);
    set_header_footer_image(options->image_center_, image_position_t::HEADER_CENTER);
    set_header_footer_image(options->image_right_, image_position_t::HEADER_RIGHT);
  }

  header_                = tmp_header;
  header_footer_changed_ = true;
}

// TODO Quite similar to set_header. Maybe merged in one generic function
void worksheet_t::set_footer(const std::string& str, const std::optional<header_footer_options_t>& options)
{
  ///     lxw_error err;
  ///     char *tmp_footer;
  ///     char *found_string;
  ///     char *offset_string;
  ///     uint8_t placeholder_count = 0;
  ///     uint8_t image_count = 0;
  ///
  if(str.empty())
  {
    throw xwpp_exception_t("Footer is empty");
  }

  if(str.size() > HEADER_FOOTER_MAX)
  {
    throw xwpp_out_of_range_t(std::format("Footer {} is empty too long ({}/{})", str, str.size(), HEADER_FOOTER_MAX));
  }

  std::string tmp_footer = str;

  // Count &G placeholders and ensure there are sufficient images.
  uint8_t placeholder_count = 0;
  for(size_t i = 0; i < tmp_footer.size() - 1; ++i)
  {
    if(tmp_footer[i] == '&' && tmp_footer[i + 1] == 'G')
    {
      placeholder_count++;
    }
  }

  if(placeholder_count > 0 && !options)
  {
    throw xwpp_exception_t("&G placeholders present but no image provided");
  }

  if(options)
  {
    uint8_t image_count = 0;

    /* Ensure there are enough images to match the placeholders. There is
     * a potential bug where there are sufficient images but in the wrong
     * positions but we don't currently try to deal with that.*/
    if(!options->image_left_.empty())
    {
      image_count++;
    }
    if(!options->image_center_.empty())
    {
      image_count++;
    }
    if(!options->image_right_.empty())
    {
      image_count++;
    }

    if(placeholder_count != image_count)
    {
      throw xwpp_exception_t("Number of &G placeholders does not match number of supplied images");
    }

    if(options->margin_ > 0.0)
    {
      margin_footer_ = options->margin_;
    }

    set_header_footer_image(options->image_left_, image_position_t::FOOTER_LEFT);
    set_header_footer_image(options->image_center_, image_position_t::FOOTER_CENTER);
    set_header_footer_image(options->image_right_, image_position_t::FOOTER_RIGHT);
  }

  footer_                = tmp_footer;
  header_footer_changed_ = true;
}

void worksheet_t::set_header(const std::string& str)
{
  set_header(str, std::nullopt);
}

void worksheet_t::set_footer(const std::string& str)
{
  set_footer(str, std::nullopt);
}

/// void
/// worksheet_gridlines(lxw_worksheet *self, uint8_t option)
/// {
///     if (option == LXW_HIDE_ALL_GRIDLINES) {
///         self->print_gridlines = 0;
///         self->screen_gridlines = 0;
///     }
///
///     if (option & LXW_SHOW_SCREEN_GRIDLINES) {
///         self->screen_gridlines = 1;
///     }
///
///     if (option & LXW_SHOW_PRINT_GRIDLINES) {
///         self->print_gridlines = 1;
///         self->print_options_changed = 1;
///     }
/// }

/// void
/// worksheet_center_horizontally(lxw_worksheet *self)
/// {
///     self->print_options_changed = 1;
///     self->hcenter = 1;
/// }

/// void
/// worksheet_center_vertically(lxw_worksheet *self)
/// {
///     self->print_options_changed = 1;
///     self->vcenter = 1;
/// }

/// void
/// worksheet_print_row_col_headers(lxw_worksheet *self)
/// {
///     self->print_headers = 1;
///     self->print_options_changed = 1;
/// }

/// lxw_error
/// worksheet_repeat_rows(lxw_worksheet *self, row_num_t first_row,
///                       row_num_t last_row)
/// {
///     row_num_t tmp_row;
///     lxw_error err;
///
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///
///     err = _check_dimensions(self, last_row, 0, LXW_IGNORE, LXW_IGNORE);
///     if (err)
///         return err;
///
///     self->repeat_rows.in_use = LXW_TRUE;
///     self->repeat_rows.first_row = first_row;
///     self->repeat_rows.last_row = last_row;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_repeat_columns(lxw_worksheet *self, col_num_t first_col,
///                          col_num_t last_col)
/// {
///     col_num_t tmp_col;
///     lxw_error err;
///
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     err = _check_dimensions(self, last_col, 0, LXW_IGNORE, LXW_IGNORE);
///     if (err)
///         return err;
///
///     self->repeat_cols.in_use = LXW_TRUE;
///     self->repeat_cols.first_col = first_col;
///     self->repeat_cols.last_col = last_col;
///
///     return LXW_NO_ERROR;
/// }

/// lxw_error
/// worksheet_print_area(lxw_worksheet *self, row_num_t first_row,
///                      col_num_t first_col, row_num_t last_row,
///                      col_num_t last_col)
/// {
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     lxw_error err;
///
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     err = _check_dimensions(self, last_row, last_col, LXW_IGNORE, LXW_IGNORE);
///     if (err)
///         return err;
///
///     /* Ignore max area since it is the same as no print area in Excel. */
///     if (first_row == 0 && first_col == 0 && last_row == LXW_ROW_MAX - 1
///         && last_col == LXW_COL_MAX - 1) {
///         return LXW_NO_ERROR;
///     }
///
///     self->print_area.in_use = LXW_TRUE;
///     self->print_area.first_row = first_row;
///     self->print_area.last_row = last_row;
///     self->print_area.first_col = first_col;
///     self->print_area.last_col = last_col;
///
///     return LXW_NO_ERROR;
/// }

/// void
/// worksheet_fit_to_pages(lxw_worksheet *self, uint16_t width, uint16_t height)
/// {
///     self->fit_page = 1;
///     self->fit_width = width;
///     self->fit_height = height;
///     self->page_setup_changed = 1;
/// }

/// void
/// worksheet_set_start_page(lxw_worksheet *self, uint16_t start_page)
/// {
///     self->page_start = start_page;
/// }

/// void
/// worksheet_set_print_scale(lxw_worksheet *self, uint16_t scale)
/// {
///     /* Confine the scale to Excel"s range */
///     if (scale < 10 || scale > 400)
///         return;
///
///     /* Turn off "fit to page" option. */
///     self->fit_page = LXW_FALSE;
///
///     self->print_scale = scale;
///     self->page_setup_changed = LXW_TRUE;
/// }

/// void
/// worksheet_print_black_and_white(lxw_worksheet *self)
/// {
///     self->black_white = LXW_TRUE;
///     self->page_setup_changed = LXW_TRUE;
/// }

void worksheet_t::set_h_pagebreaks(const std::vector<row_num_t>& breaks)
{
  hbreaks_ = breaks;

  /* The Excel 2007 specification says that the maximum number of page
   * breaks is 1026. However, in practice it is actually 1023. */
  if(hbreaks_.size() > BREAKS_MAX)
  {
    hbreaks_.resize(BREAKS_MAX);
  }
}

void worksheet_t::set_v_pagebreaks(const std::vector<col_num_t>& breaks)
{
  vbreaks_ = breaks;

  /* The Excel 2007 specification says that the maximum number of page
   * breaks is 1026. However, in practice it is actually 1023. */
  if(vbreaks_.size() > BREAKS_MAX)
  {
    vbreaks_.resize(BREAKS_MAX);
  }
}

/// void
/// worksheet_set_zoom(lxw_worksheet *self, uint16_t scale)
/// {
///     /* Confine the scale to Excel"s range */
///     if (scale < 10 || scale > 400) {
///         LXW_WARN("worksheet_set_zoom(): "
///                  "Zoom factor scale outside range: 10 <= zoom <= 400.");
///         return;
///     }
///
///     self->zoom = scale;
/// }

/// void
/// worksheet_hide_zero(lxw_worksheet *self)
/// {
///     self->show_zeros = LXW_FALSE;
/// }

/// void
/// worksheet_right_to_left(lxw_worksheet *self)
/// {
///     self->right_to_left = LXW_TRUE;
/// }

void worksheet_t::set_tab_color(color_t color)
{
  tab_color_ = color;
}

void worksheet_t::protect(const std::string& password)
{
  protect(password, std::nullopt);
}

void worksheet_t::protect(std::optional<protection_t> options)
{
  protect("", options);
}

void worksheet_t::protect()
{
  protect("", std::nullopt);
}

void worksheet_t::protect(const std::string& password, std::optional<protection_t> options)
{
  // Copy any user parameters to the internal structure.
  if(options)
  {
    protection_.no_select_locked_cells_   = options->no_select_locked_cells_;
    protection_.no_select_unlocked_cells_ = options->no_select_unlocked_cells_;
    protection_.format_cells_             = options->format_cells_;
    protection_.format_columns_           = options->format_columns_;
    protection_.format_rows_              = options->format_rows_;
    protection_.insert_columns_           = options->insert_columns_;
    protection_.insert_rows_              = options->insert_rows_;
    protection_.insert_hyperlinks_        = options->insert_hyperlinks_;
    protection_.delete_columns_           = options->delete_columns_;
    protection_.delete_rows_              = options->delete_rows_;
    protection_.sort_                     = options->sort_;
    protection_.autofilter_               = options->autofilter_;
    protection_.pivot_tables_             = options->pivot_tables_;
    protection_.scenarios_                = options->scenarios_;
    protection_.objects_                  = options->objects_;
  }

  if(!password.empty())
  {
    const uint16_t hash = hash_password(password);
    protection_.hash_   = std::format("{:5X}", hash);
  }

  protection_.no_sheet_      = false;
  protection_.no_content_    = true;
  protection_.is_configured_ = true;
}

/// void
/// worksheet_outline_settings(lxw_worksheet *self,
///                            uint8_t visible,
///                            uint8_t symbols_below,
///                            uint8_t symbols_right, uint8_t auto_style)
/// {
///     self->outline_on = visible;
///     self->outline_below = symbols_below;
///     self->outline_right = symbols_right;
///     self->outline_style = auto_style;
///
///     self->outline_changed = LXW_TRUE;
/// }

/// void
/// worksheet_set_default_row(lxw_worksheet *self, double height,
///                           uint8_t hide_unused_rows)
/// {
///     if (height < 0)
///         height = self->default_row_height;
///
///     if (height != self->default_row_height) {
///         self->default_row_height = height;
///         self->row_size_changed = LXW_TRUE;
///     }
///
///     if (hide_unused_rows)
///         self->default_row_zeroed = LXW_TRUE;
///
///     self->default_row_set = LXW_TRUE;
/// }

void worksheet_t::insert_image(row_num_t row_num, col_num_t col_num, const std::string& filename,
                               std::optional<image_options_t> user_options)
{
  if(filename.empty())
  {
    throw xwpp_exception_t("Image filename is empty");
  }

  // Check that the image file exists and can be opened.
  {
    std::ifstream image_stream(filename);
    if(!image_stream)
    {
      throw xwpp_exception_t(std::format("Image file {} doesn't exist or cannot be opened", filename));
    }
  }

  // Use the filename as the default description, like Excel.
  // TODO Use basename of file, not fullname
  std::string description = filename;

  // Create a new object to hold the image properties.
  object_properties_t object_props;

  if(user_options)
  {
    object_props.x_offset_        = user_options->x_offset_;
    object_props.y_offset_        = user_options->y_offset_;
    object_props.x_scale_         = user_options->x_scale_;
    object_props.y_scale_         = user_options->y_scale_;
    object_props.url_             = user_options->url_;
    object_props.tip_             = user_options->tip_;
    object_props.object_position_ = user_options->object_position_;
    object_props.decorative_      = user_options->decorative_;

    if(!user_options->description_.empty())
    {
      description = user_options->description_;
    }
  }

  // Copy other options or set defaults.
  object_props.filename_    = filename;
  object_props.description_ = description;
  object_props.row_         = row_num;
  object_props.col_         = col_num;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  get_image_properties(object_props);
  image_props_.push_back(object_props);
}

void worksheet_t::insert_image(row_num_t row_num, col_num_t col_num, const std::string& filename)
{
  insert_image(row_num, col_num, filename, std::nullopt);
}

void worksheet_t::insert_image_buffer(row_num_t row_num, col_num_t col_num,
                                      const std::vector<unsigned char>& image_buffer,
                                      std::optional<image_options_t> user_options)
{
  if(image_buffer.empty())
  {
    throw xwpp_exception_t("Image is empty");
  }

  object_properties_t object_props;

  object_props.image_buffer_ = image_buffer;
  if(user_options)
  {
    object_props.x_offset_        = user_options->x_offset_;
    object_props.y_offset_        = user_options->y_offset_;
    object_props.x_scale_         = user_options->x_scale_;
    object_props.y_scale_         = user_options->y_scale_;
    object_props.url_             = user_options->url_;
    object_props.tip_             = user_options->tip_;
    object_props.object_position_ = user_options->object_position_;
    object_props.description_     = user_options->description_;
    object_props.decorative_      = user_options->decorative_;
  }

  // Copy other options or set defaults.
  object_props.filename_ = "image_buffer";
  object_props.row_      = row_num;
  object_props.col_      = col_num;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  get_image_properties(object_props);
  image_props_.push_back(object_props);
}

void worksheet_t::insert_image_buffer(row_num_t row_num, col_num_t col_num,
                                      const std::vector<unsigned char>& image_buffer)
{
  insert_image_buffer(row_num, col_num, image_buffer, std::nullopt);
}

void worksheet_t::embed_image(row_num_t row_num, col_num_t col_num, const std::string& filename,
                              std::optional<image_options_t> options)
{
  if(filename.empty())
  {
    throw xwpp_exception_t("Image filename is empty");
  }

  // Check that the image file exists and can be opened.
  {
    std::ifstream image_stream(filename);
    if(!image_stream)
    {
      throw xwpp_exception_t(std::format("Image file {} doesn't exist or cannot be opened", filename));
    }
  }

  // Check and store the cell dimensions.
  check_dimensions(row_num, col_num, false, false);

  // Create a new object to hold the image properties.
  object_properties_t object_props;

  // We only copy/use a limited number of options for embedded images.
  if(options)
  {
    if(options->cell_format_)
    {
      object_props.format_ = options->cell_format_;
    }

    // The url for embedded images is written as a cell url.
    if(!options->url_.empty())
    {
      if(!options->cell_format_)
      {
        object_props.format_ = default_url_format_;
      }

      // TODO Don't use a member data but a local variable
      storing_embedded_image_ = true;
      write_url(row_num, col_num, options->url_, object_props.format_);
      storing_embedded_image_ = false;
    }

    object_props.decorative_ = options->decorative_;
    if(!options->description_.empty())
    {
      object_props.description_ = options->description_;
    }
  }
  // Copy other options or set defaults.
  object_props.filename_ = filename;
  object_props.row_      = row_num;
  object_props.col_      = col_num;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  get_image_properties(object_props);
  embedded_image_props_.push_back(object_props);
}

void worksheet_t::embed_image(row_num_t row_num, col_num_t col_num, const std::string& filename)
{
  embed_image(row_num, col_num, filename, std::nullopt);
}

void worksheet_t::embed_image_buffer(row_num_t row_num, col_num_t col_num,
                                     const std::vector<unsigned char>& image_buffer,
                                     std::optional<image_options_t> options)
{
  if(image_buffer.empty())
  {
    throw xwpp_exception_t("Image is empty");
  }

  check_dimensions(row_num, col_num, false, false);

  object_properties_t object_props;
  object_props.image_buffer_ = image_buffer;

  // We only copy/use a limited number of options for embedded images.
  if(options)
  {
    if(options->cell_format_)
    {
      object_props.format_ = options->cell_format_;
    }

    // The url for embedded images is written as a cell url.
    if(!options->url_.empty())
    {
      if(!options->cell_format_)
      {
        object_props.format_ = default_url_format_;
      }

      // TODO Don't use a member data but a local variable
      storing_embedded_image_ = true;
      write_url(row_num, col_num, options->url_, object_props.format_);
      storing_embedded_image_ = false;
    }

    object_props.decorative_ = options->decorative_;
    if(!options->description_.empty())
    {
      object_props.description_ = options->description_;
    }
  }
  // Copy other options or set defaults.
  object_props.filename_ = "image_buffer";
  object_props.row_      = row_num;
  object_props.col_      = col_num;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  get_image_properties(object_props);
  embedded_image_props_.push_back(object_props);
}

void worksheet_t::embed_image_buffer(row_num_t row_num, col_num_t col_num,
                                     const std::vector<unsigned char>& image_buffer)
{
  embed_image_buffer(row_num, col_num, image_buffer, std::nullopt);
}

/// lxw_error
/// worksheet_set_background(lxw_worksheet *self, const char *filename)
/// {
///     FILE *image_stream;
///     lxw_object_properties *object_props;
///
///     if (!filename) {
///         LXW_WARN("worksheet_set_background(): "
///                  "filename must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }
///
///     /* Check that the image file exists and can be opened. */
///     image_stream = lxw_fopen(filename, "rb");
///     if (!image_stream) {
///         LXW_WARN_FORMAT1("worksheet_set_background(): "
///                          "file doesn't exist or can't be opened: %s.",
///                          filename);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Create a new object to hold the image properties. */
///     object_props = calloc(1, sizeof(lxw_object_properties));
///     if (!object_props) {
///         fclose(image_stream);
///         return LXW_ERROR_MEMORY_MALLOC_FAILED;
///     }
///
///     /* Copy other options or set defaults. */
///     object_props->filename = lxw_strdup(filename);
///     object_props->stream = image_stream;
///     object_props->is_background = LXW_TRUE;
///
///     if (_get_image_properties(object_props) == LXW_NO_ERROR) {
///         _free_object_properties(self->background_image);
///         self->background_image = object_props;
///         self->has_background_image = LXW_TRUE;
///         fclose(image_stream);
///         return LXW_NO_ERROR;
///     }
///     else {
///         _free_object_properties(object_props);
///         fclose(image_stream);
///         return LXW_ERROR_IMAGE_DIMENSIONS;
///     }
/// }

/// lxw_error
/// worksheet_set_background_buffer(lxw_worksheet *self,
///                                 const unsigned char *image_buffer,
///                                 size_t image_size)
/// {
///     FILE *image_stream;
///     lxw_object_properties *object_props;
///
///     if (!image_size) {
///         LXW_WARN("worksheet_set_background(): " "size must be non-zero.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }
///
///     /* Write the image buffer to a file (preferably in memory) so we can read
///      * the dimensions like an ordinary file. */
/// #ifdef USE_FMEMOPEN
///     image_stream = fmemopen((void *) image_buffer, image_size, "rb");
///
///     if (!image_stream)
///         return LXW_ERROR_CREATING_TMPFILE;
/// #else
///     image_stream = lxw_tmpfile(self->tmpdir);
///
///     if (!image_stream)
///         return LXW_ERROR_CREATING_TMPFILE;
///
///     if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
///         fclose(image_stream);
///         return LXW_ERROR_CREATING_TMPFILE;
///     }
///
///     rewind(image_stream);
/// #endif
///
///     /* Create a new object to hold the image properties. */
///     object_props = calloc(1, sizeof(lxw_object_properties));
///     if (!object_props) {
///         fclose(image_stream);
///         return LXW_ERROR_MEMORY_MALLOC_FAILED;
///     }
///
///     /* Store the image data in the properties structure. */
///     object_props->image_buffer = calloc(1, image_size);
///     if (!object_props->image_buffer) {
///         _free_object_properties(object_props);
///         fclose(image_stream);
///         return LXW_ERROR_MEMORY_MALLOC_FAILED;
///     }
///     else {
///         memcpy(object_props->image_buffer, image_buffer, image_size);
///         object_props->image_buffer_size = image_size;
///         object_props->is_image_buffer = LXW_TRUE;
///     }
///
///     /* Copy other options or set defaults. */
///     object_props->filename = lxw_strdup("image_buffer");
///     object_props->stream = image_stream;
///     object_props->is_background = LXW_TRUE;
///
///     if (_get_image_properties(object_props) == LXW_NO_ERROR) {
///         _free_object_properties(self->background_image);
///         self->background_image = object_props;
///         self->has_background_image = LXW_TRUE;
///         fclose(image_stream);
///         return LXW_NO_ERROR;
///     }
///     else {
///         _free_object_properties(object_props);
///         fclose(image_stream);
///         return LXW_ERROR_IMAGE_DIMENSIONS;
///     }
/// }

void worksheet_t::insert_chart(row_num_t row_num, col_num_t col_num, chart_t* chart,
                               const std::optional<chart_options_t>& user_options)
{
  if(chart->in_use_)
  {
    throw xwpp_exception_t("The same chart object cannot be inserted in a worksheet more than once");
  }

  if(chart->series_list_.empty())
  {
    throw xwpp_exception_t("Chart must have a series");
  }

  object_properties_t object_props;
  if(user_options)
  {
    object_props.x_offset_        = user_options->x_offset_;
    object_props.y_offset_        = user_options->y_offset_;
    object_props.x_scale_         = user_options->x_scale_;
    object_props.y_scale_         = user_options->y_scale_;
    object_props.object_position_ = user_options->object_position_;
    object_props.description_     = user_options->description_;
    object_props.decorative_      = user_options->decorative_;
  }

  // Copy other options or set defaults.
  object_props.row_ = row_num;
  object_props.col_ = col_num;

  object_props.width_  = 480;
  object_props.height_ = 288;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  // Store chart references so they can be ordered in the workbook.
  object_props.chart_ = chart;

  chart_data_.push_back(object_props);

  chart->in_use_ = true;
}

void worksheet_t::insert_chart(row_num_t row_num, col_num_t col_num, chart_t* chart)
{
  return insert_chart(row_num, col_num, chart, std::nullopt);
}

/// lxw_error
/// worksheet_data_validation_range(lxw_worksheet *self, row_num_t first_row,
///                                 col_num_t first_col,
///                                 row_num_t last_row,
///                                 col_num_t last_col,
///                                 lxw_data_validation *validation)
/// {
///     lxw_data_val_obj *copy;
///     uint8_t is_between = LXW_FALSE;
///     uint8_t is_formula = LXW_FALSE;
///     uint8_t has_criteria = LXW_TRUE;
///     lxw_error err;
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     size_t length;
///
///     /* No action is required for validation type 'any' unless there are
///      * input messages to display.*/
///     if (validation->validate == LXW_VALIDATION_TYPE_ANY
///         && !(validation->input_title || validation->input_message)) {
///
///         return LXW_NO_ERROR;
///     }
///
///     /* Check for formula types. */
///     switch (validation->validate) {
///         case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
///         case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
///         case LXW_VALIDATION_TYPE_LIST_FORMULA:
///         case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
///         case LXW_VALIDATION_TYPE_DATE_FORMULA:
///         case LXW_VALIDATION_TYPE_TIME_FORMULA:
///         case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
///             is_formula = LXW_TRUE;
///             break;
///     }
///
///     /* Check for types without a criteria. */
///     switch (validation->validate) {
///         case LXW_VALIDATION_TYPE_LIST:
///         case LXW_VALIDATION_TYPE_LIST_FORMULA:
///         case LXW_VALIDATION_TYPE_ANY:
///         case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
///             has_criteria = LXW_FALSE;
///             break;
///     }
///
///     /* Check that a validation parameter has been specified
///      * except for 'list', 'any' and 'custom'. */
///     if (has_criteria && validation->criteria == LXW_VALIDATION_CRITERIA_NONE) {
///
///         LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
///                         "criteria parameter must be specified.");
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Check for "between" criteria so we can do additional checks. */
///     if (has_criteria
///         && (validation->criteria == LXW_VALIDATION_CRITERIA_BETWEEN
///             || validation->criteria == LXW_VALIDATION_CRITERIA_NOT_BETWEEN)) {
///
///         is_between = LXW_TRUE;
///     }
///
///     /* Check that formula values are non NULL. */
///     if (is_formula) {
///         if (is_between) {
///             if (!validation->minimum_formula) {
///                 LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
///                                 "minimum_formula parameter cannot be NULL.");
///                 return LXW_ERROR_PARAMETER_VALIDATION;
///             }
///             if (!validation->maximum_formula) {
///                 LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
///                                 "maximum_formula parameter cannot be NULL.");
///                 return LXW_ERROR_PARAMETER_VALIDATION;
///             }
///         }
///         else {
///             if (!validation->value_formula) {
///                 LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
///                                 "formula parameter cannot be NULL.");
///                 return LXW_ERROR_PARAMETER_VALIDATION;
///             }
///         }
///     }
///
///     /* Check Excel limitations on input strings. */
///     if (validation->input_title) {
///         length = lxw_utf8_strlen(validation->input_title);
///         if (length > LXW_VALIDATION_MAX_TITLE_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
///                              "input_title length > Excel limit of %d.",
///                              LXW_VALIDATION_MAX_TITLE_LENGTH);
///             return LXW_ERROR_32_STRING_LENGTH_EXCEEDED;
///         }
///     }
///
///     if (validation->error_title) {
///         length = lxw_utf8_strlen(validation->error_title);
///         if (length > LXW_VALIDATION_MAX_TITLE_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
///                              "error_title length > Excel limit of %d.",
///                              LXW_VALIDATION_MAX_TITLE_LENGTH);
///             return LXW_ERROR_32_STRING_LENGTH_EXCEEDED;
///         }
///     }
///
///     if (validation->input_message) {
///         length = lxw_utf8_strlen(validation->input_message);
///         if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
///                              "input_message length > Excel limit of %d.",
///                              LXW_VALIDATION_MAX_STRING_LENGTH);
///             return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///         }
///     }
///
///     if (validation->error_message) {
///         length = lxw_utf8_strlen(validation->error_message);
///         if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
///                              "error_message length > Excel limit of %d.",
///                              LXW_VALIDATION_MAX_STRING_LENGTH);
///             return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///         }
///     }
///
///     if (validation->validate == LXW_VALIDATION_TYPE_LIST) {
///         length = _validation_list_length(validation->value_list);
///
///         if (length == 0) {
///             LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
///                             "list parameters cannot be zero.");
///             return LXW_ERROR_PARAMETER_VALIDATION;
///         }
///
///         if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
///                              "list length with commas > Excel limit of %d.",
///                              LXW_VALIDATION_MAX_STRING_LENGTH);
///             return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
///         }
///     }
///
///     /* Swap last row/col with first row/col as necessary */
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     /* Check that dimensions are valid but don't store them. */
///     err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
///     if (err)
///         return err;
///
///     /* Create a copy of the parameters from the user data validation. */
///     copy = calloc(1, sizeof(lxw_data_val_obj));
///     GOTO_LABEL_ON_MEM_ERROR(copy, mem_error);
///
///     /* Create the data validation range. */
///     if (first_row == last_row && first_col == last_col)
///         lxw_rowcol_to_cell(copy->sqref, first_row, first_col);
///     else
///         lxw_rowcol_to_range(copy->sqref, first_row, first_col, last_row,
///                             last_col);
///
///     /* Copy the parameters from the user data validation. */
///     copy->validate = validation->validate;
///     copy->value_number = validation->value_number;
///     copy->error_type = validation->error_type;
///     copy->dropdown = validation->dropdown;
///
///     if (has_criteria)
///         copy->criteria = validation->criteria;
///
///     if (is_between) {
///         copy->value_number = validation->minimum_number;
///         copy->maximum_number = validation->maximum_number;
///     }
///
///     /* Copy the input/error titles and messages. */
///     if (validation->input_title) {
///         copy->input_title = lxw_strdup_formula(validation->input_title);
///         GOTO_LABEL_ON_MEM_ERROR(copy->input_title, mem_error);
///     }
///
///     if (validation->input_message) {
///         copy->input_message = lxw_strdup_formula(validation->input_message);
///         GOTO_LABEL_ON_MEM_ERROR(copy->input_message, mem_error);
///     }
///
///     if (validation->error_title) {
///         copy->error_title = lxw_strdup_formula(validation->error_title);
///         GOTO_LABEL_ON_MEM_ERROR(copy->error_title, mem_error);
///     }
///
///     if (validation->error_message) {
///         copy->error_message = lxw_strdup_formula(validation->error_message);
///         GOTO_LABEL_ON_MEM_ERROR(copy->error_message, mem_error);
///     }
///
///     /* Copy the formula strings. */
///     if (is_formula) {
///         if (is_between) {
///             copy->value_formula =
///                 lxw_strdup_formula(validation->minimum_formula);
///             GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
///             copy->maximum_formula =
///                 lxw_strdup_formula(validation->maximum_formula);
///             GOTO_LABEL_ON_MEM_ERROR(copy->maximum_formula, mem_error);
///         }
///         else {
///             copy->value_formula =
///                 lxw_strdup_formula(validation->value_formula);
///             GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
///         }
///     }
///
///     /* Copy the validation list as a csv string. */
///     if (validation->validate == LXW_VALIDATION_TYPE_LIST) {
///         copy->value_formula = _validation_list_to_csv(validation->value_list);
///         GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
///     }
///
///     if (validation->validate == LXW_VALIDATION_TYPE_DATE
///         || validation->validate == LXW_VALIDATION_TYPE_TIME) {
///         if (is_between) {
///             copy->value_number =
///                 lxw_datetime_to_excel_date_with_epoch
///                 (&validation->minimum_datetime, self->use_1904_epoch);
///             copy->maximum_number =
///                 lxw_datetime_to_excel_date_with_epoch
///                 (&validation->maximum_datetime, self->use_1904_epoch);
///         }
///         else {
///             copy->value_number =
///                 lxw_datetime_to_excel_date_with_epoch
///                 (&validation->value_datetime, self->use_1904_epoch);
///         }
///     }
///
///     /* These options are on by default so we can't take plain booleans. */
///     copy->ignore_blank = validation->ignore_blank ^ 1;
///     copy->show_input = validation->show_input ^ 1;
///     copy->show_error = validation->show_error ^ 1;
///
///     STAILQ_INSERT_TAIL(self->data_validations, copy, list_pointers);
///
///     self->num_validations++;
///
///     return LXW_NO_ERROR;
///
/// mem_error:
///     _free_data_validation(copy);
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
/// }

/// lxw_error
/// worksheet_data_validation_cell(lxw_worksheet *self, row_num_t row,
///                                col_num_t col, lxw_data_validation *validation)
/// {
///     return worksheet_data_validation_range(self, row, col,
///                                            row, col, validation);
/// }

/// lxw_error
/// worksheet_conditional_format_range(lxw_worksheet *self, row_num_t first_row,
///                                    col_num_t first_col,
///                                    row_num_t last_row,
///                                    col_num_t last_col,
///                                    lxw_conditional_format *user_options)
/// {
///     lxw_cond_format_obj *cond_format;
///     row_num_t tmp_row;
///     col_num_t tmp_col;
///     lxw_error err = LXW_NO_ERROR;
///     char *type_strings[] = {
///         "none",
///         "cellIs",
///         "containsText",
///         "timePeriod",
///         "aboveAverage",
///         "duplicateValues",
///         "uniqueValues",
///         "top10",
///         "top10",
///         "containsBlanks",
///         "notContainsBlanks",
///         "containsErrors",
///         "notContainsErrors",
///         "expression",
///         "colorScale",
///         "colorScale",
///         "dataBar",
///         "iconSet",
///     };
///
///     /* Swap last row/col with first row/col as necessary */
///     if (first_row > last_row) {
///         tmp_row = last_row;
///         last_row = first_row;
///         first_row = tmp_row;
///     }
///     if (first_col > last_col) {
///         tmp_col = last_col;
///         last_col = first_col;
///         first_col = tmp_col;
///     }
///
///     /* Check that dimensions are valid but don't store them. */
///     err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
///     if (err)
///         return err;
///
///     /* Check the validation type is in correct enum range. */
///     if (user_options->type <= LXW_CONDITIONAL_TYPE_NONE ||
///         user_options->type >= LXW_CONDITIONAL_TYPE_LAST) {
///
///         LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                          "invalid type value (%d).", user_options->type);
///
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///
///     /* Create a copy of the parameters from the user data validation. */
///     cond_format = calloc(1, sizeof(lxw_cond_format_obj));
///     GOTO_LABEL_ON_MEM_ERROR(cond_format, error);
///
///     /* Create the data validation range. */
///     if (first_row == last_row && first_col == last_col)
///         lxw_rowcol_to_cell(cond_format->sqref, first_row, first_col);
///     else
///         lxw_rowcol_to_range(cond_format->sqref, first_row, first_col,
///                             last_row, last_col);
///
///     /* Store the first cell string for text and date rules. */
///     lxw_rowcol_to_cell(cond_format->first_cell, first_row, first_col);
///
///     /* Overwrite the sqref range with a user supplied set of ranges. */
///     if (user_options->multi_range) {
///
///         if (strlen(user_options->multi_range) >= LXW_MAX_ATTRIBUTE_LENGTH) {
///             LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
///                              "multi_range >= limit = %d.",
///                              LXW_MAX_ATTRIBUTE_LENGTH);
///             err = LXW_ERROR_PARAMETER_VALIDATION;
///             goto error;
///         }
///
///         LXW_ATTRIBUTE_COPY(cond_format->sqref, user_options->multi_range);
///     }
///
///     /* Get the conditional format dxf format index. */
///     if (user_options->format)
///         cond_format->dxf_index =
///             lxw_format_get_dxf_index(user_options->format);
///     else
///         cond_format->dxf_index = LXW_PROPERTY_UNSET;
///
///     /* Set some common option for all validation types. */
///     cond_format->type = user_options->type;
///     cond_format->criteria = user_options->criteria;
///     cond_format->stop_if_true = user_options->stop_if_true;
///     cond_format->type_string = lxw_strdup(type_strings[cond_format->type]);
///
///     /* Check that the criteria matches the conditional type. */
///     err = _validate_conditional_criteria(cond_format);
///     if (err)
///         goto error;
///
///     /* Validate the user input for various types of rules. */
///     if (user_options->type == LXW_CONDITIONAL_TYPE_CELL
///         || cond_format->type == LXW_CONDITIONAL_TYPE_DUPLICATE
///         || cond_format->type == LXW_CONDITIONAL_TYPE_UNIQUE) {
///
///         err = _validate_conditional_cell(cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (user_options->type == LXW_CONDITIONAL_TYPE_TEXT) {
///
///         err = _validate_conditional_text(cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (user_options->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {
///
///         err = _validate_conditional_time_period(user_options);
///         if (err)
///             goto error;
///     }
///     else if (user_options->type == LXW_CONDITIONAL_TYPE_AVERAGE) {
///
///         err = _validate_conditional_average(user_options);
///         if (err)
///             goto error;
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
///              || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {
///
///         err = _validate_conditional_top(cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (user_options->type == LXW_CONDITIONAL_TYPE_FORMULA) {
///
///         err = _validate_conditional_formula(cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_2_COLOR_SCALE
///              || cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
///
///         err = _validate_conditional_scale(cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_DATA_BAR) {
///
///         err = _validate_conditional_data_bar(self, cond_format, user_options);
///         if (err)
///             goto error;
///     }
///     else if (cond_format->type == LXW_CONDITIONAL_TYPE_ICON_SETS) {
///
///         err = _validate_conditional_icons(user_options);
///         if (err)
///             goto error;
///
///         cond_format->icon_style = user_options->icon_style;
///         cond_format->reverse_icons = user_options->reverse_icons;
///         cond_format->icons_only = user_options->icons_only;
///     }
///
///     /* Set the priority based on the order of adding. */
///     cond_format->dxf_priority = ++self->dxf_priority;
///
///     /* Store the conditional format object. */
///     err = _store_conditional_format_object(self, cond_format);
///
///     if (err)
///         goto error;
///     else
///         return LXW_NO_ERROR;
///
/// error:
///     _free_cond_format(cond_format);
///     return err;
/// }

/// lxw_error
/// worksheet_conditional_format_cell(lxw_worksheet *self,
///                                   row_num_t row,
///                                   col_num_t col,
///                                   lxw_conditional_format *options)
/// {
///     return worksheet_conditional_format_range(self, row, col,
///                                               row, col, options);
/// }

/// lxw_error
/// worksheet_insert_button(lxw_worksheet *self, row_num_t row_num,
///                         col_num_t col_num, lxw_button_options *options)
/// {
///     lxw_error err;
///     lxw_vml_obj *button;
///
///     err = _check_dimensions(self, row_num, col_num, LXW_TRUE, LXW_TRUE);
///     if (err)
///         return err;
///
///     button = calloc(1, sizeof(lxw_vml_obj));
///     GOTO_LABEL_ON_MEM_ERROR(button, mem_error);
///
///     button->row = row_num;
///     button->col = col_num;
///
///     /* Set user and default parameters for the button. */
///     err = _get_button_params(button, 1 + self->num_buttons, options);
///     if (err)
///         goto mem_error;
///
///     /* Calculate the worksheet position of the button. */
///     _worksheet_position_vml_object(self, button);
///
///     self->has_vml = LXW_TRUE;
///     self->has_buttons = LXW_TRUE;
///     self->num_buttons++;
///
///     STAILQ_INSERT_TAIL(self->button_objs, button, list_pointers);
///
///     return LXW_NO_ERROR;
///
/// mem_error:
///     if (button)
///         _free_vml_object(button);
///
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
/// }

/// lxw_error
/// worksheet_set_vba_name(lxw_worksheet *self, const char *name)
/// {
///     if (!name) {
///         LXW_WARN("worksheet_set_vba_name(): " "name must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }
///
///     self->vba_codename = lxw_strdup(name);
///
///     return LXW_NO_ERROR;
/// }

/// void
/// worksheet_set_comments_author(lxw_worksheet *self, const char *author)
/// {
///     self->comment_author = lxw_strdup(author);
/// }

void worksheet_t::show_comments()
{
  comment_display_default_ = comment_display_t::VISIBLE;
}

/// lxw_error
/// worksheet_ignore_errors(lxw_worksheet *self, uint8_t type, const char *range)
/// {
///     if (!range) {
///         LXW_WARN("worksheet_ignore_errors(): " "'range' must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }
///
///     if (type <= 0 || type >= LXW_IGNORE_LAST_OPTION) {
///         LXW_WARN("worksheet_ignore_errors(): " "unknown option in 'type'.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }
///
///     /* Set the ranges to be ignored. */
///     if (type == LXW_IGNORE_NUMBER_STORED_AS_TEXT) {
///         free(self->ignore_number_stored_as_text);
///         self->ignore_number_stored_as_text = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_EVAL_ERROR) {
///         free(self->ignore_eval_error);
///         self->ignore_eval_error = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_FORMULA_DIFFERS) {
///         free(self->ignore_formula_differs);
///         self->ignore_formula_differs = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_FORMULA_RANGE) {
///         free(self->ignore_formula_range);
///         self->ignore_formula_range = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_FORMULA_UNLOCKED) {
///         free(self->ignore_formula_unlocked);
///         self->ignore_formula_unlocked = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_EMPTY_CELL_REFERENCE) {
///         free(self->ignore_empty_cell_reference);
///         self->ignore_empty_cell_reference = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_LIST_DATA_VALIDATION) {
///         free(self->ignore_list_data_validation);
///         self->ignore_list_data_validation = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_CALCULATED_COLUMN) {
///         free(self->ignore_calculated_column);
///         self->ignore_calculated_column = lxw_strdup(range);
///     }
///     else if (type == LXW_IGNORE_TWO_DIGIT_TEXT_YEAR) {
///         free(self->ignore_two_digit_text_year);
///         self->ignore_two_digit_text_year = lxw_strdup(range);
///     }
///
///     self->has_ignore_errors = LXW_TRUE;
///
///     return LXW_NO_ERROR;
/// }

void worksheet_t::set_error_cell(const object_properties_t& object_props, uint32_t ref_id)
{
  row_num_t row_num = object_props.row_;
  col_num_t col_num = object_props.col_;

  cell_t cell = new_error_cell(row_num, col_num, ref_id, object_props.format_);
  insert_cell(row_num, col_num, cell);
}

}
