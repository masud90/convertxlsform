#' Convert an XLSForm to a Word document
#'
#' @importFrom readxl    excel_sheets read_excel
#' @importFrom officer   read_docx body_add_par body_add_fpar body_add_img
#' @importFrom officer   fpar ftext fp_text fp_par fp_border block_list external_img body_add_blocks
#' @importFrom officer   body_add_flextable
#' @importFrom commonmark markdown_html
#' @importFrom flextable regulartable border_outer delete_part width body_add_flextable autofit align compose as_paragraph valign add_header_lines colformat_md
#' @importFrom magick    image_read image_info
#' @importFrom stats setNames
#' @param xlsform_path Path to the .xlsx XLSForm
#' @param selected_language Language code (defaults to "en")
#' @param output_docx Destination .docx file (defaults to basename(xlsx))
#' @param number_questions Logical: number questions? (defaults to TRUE)
#' @param choice_code Logical: if TRUE, include choice code (name) in square brackets; otherwise show only label. (defaults to FALSE)
#' @param media_dir Optional path to the media folder (default: directory of xlsform or \code{file.path(dirname(xlsform_path), "media")} if it exists)
#' @return Invisibly returns the path to the generated .docx
#' @examples
#' \dontrun{
#' convertxlsform("survey.xlsx")
#' convertxlsform("survey.xlsx", number_questions = FALSE)
#' convertxlsform("survey.xlsx", choice_code = TRUE)
#' }
#' @export
convertxlsform <- function(xlsform_path,
                           selected_language = "en",
                           output_docx = NULL,
                           number_questions = TRUE,
                           choice_code = FALSE,
                           media_dir = NULL) {
  start_time <- Sys.time()

  # %||% : treat NULL/NA as missing
  `%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || isTRUE(is.na(a))) b else a

  # Styles (hoisted)
  fp_normal <- fp_text(font.family = "Arial", font.size = 11)
  fp_title  <- fp_text(font.family = "Arial", font.size = 16, bold = TRUE)
  fp_center <- fp_par(text.align = "center")

  # media_dir
  base_dir <- dirname(xlsform_path)
  if (is.null(media_dir)) {
    media_candidate <- file.path(base_dir, "media")
    media_dir <- if (dir.exists(media_candidate)) media_candidate else base_dir
  }

  # Read sheets
  sheets_raw   <- excel_sheets(xlsform_path)
  sheets_lower <- tolower(sheets_raw)
  survey_tbl   <- read_survey(xlsform_path, sheets_raw, sheets_lower)
  choices_tbl  <- read_choices(xlsform_path, sheets_raw, sheets_lower)
  settings_tbl <- read_settings(xlsform_path, sheets_raw, sheets_lower)
  ext_choices  <- read_external_choices(xlsform_path, sheets_raw, sheets_lower)
  q_images     <- read_images(xlsform_path, sheets_raw, sheets_lower, media_dir)

  # Validate
  validate_xlsform(
    survey_tbl, choices_tbl, settings_tbl, selected_language,
    ext_choices = ext_choices
  )

  # Parse & build
  questions    <- parse_survey_rows(survey_tbl, selected_language)
  choices_map  <- build_choices_map(choices_tbl, selected_language, media_dir,
                                    choice_code = choice_code)

  # Generate
  doc <- generate_docx(
    questions, choices_map, settings_tbl, selected_language,
    q_images, number_questions, choice_code, media_dir,
    fp_normal = fp_normal, fp_title = fp_title, fp_center = fp_center,
    ext_choices = ext_choices
  )

  # Output path
  if (is.null(output_docx)) {
    base <- tools::file_path_sans_ext(basename(xlsform_path))
    output_docx <- file.path(getwd(), paste0(base, ".docx"))
  }

  # Save
  save_docx(doc, output_docx)

  # Duration message hh:mm:ss
  secs <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))
  msg <- format_duration_message(secs)
  message(sprintf("Document generated in %s: %s", msg, normalizePath(output_docx)))

  invisible(output_docx)
}

# ---------------- Utilities: messages & errors ----------------
errf <- function(msg, ...) stop(sprintf(msg, ...), call. = FALSE)
warnf <- function(msg, ...) warning(sprintf(msg, ...), call. = FALSE)

format_duration_message <- function(secs) {
  secs <- floor(secs)
  if (secs < 60) return(sprintf("%d secs", secs))
  mins <- secs %/% 60
  rems <- secs %% 60
  if (mins < 60) return(sprintf("%d mins %d secs", mins, rems))
  hrs  <- mins %/% 60
  remm <- mins %% 60
  sprintf("%d hrs %d mins %d secs", hrs, remm, rems)
}

# ---------------- Read sheets ----------------
default_sheet_error <- function(name) {
  errf("Incomplete XLSForm: required worksheet '%s' is not present.", name)
}

read_survey <- function(path, sheets_raw, sheets_lower) {
  idx <- match("survey", sheets_lower)
  if (is.na(idx)) default_sheet_error("survey")
  read_excel(path, sheet = sheets_raw[idx])
}

read_choices <- function(path, sheets_raw, sheets_lower) {
  idx <- match("choices", sheets_lower)
  if (is.na(idx)) default_sheet_error("choices")
  read_excel(path, sheet = sheets_raw[idx])
}

read_settings <- function(path, sheets_raw, sheets_lower) {
  idx <- match("settings", sheets_lower)
  if (is.na(idx)) default_sheet_error("settings")
  read_excel(path, sheet = sheets_raw[idx])
}

read_external_choices <- function(path, sheets_raw, sheets_lower) {
  idx <- match("external_choices", sheets_lower)
  if (is.na(idx)) return(NULL)
  read_excel(path, sheet = sheets_raw[idx])
}

read_images <- function(path, sheets_raw, sheets_lower, media_dir) {
  idx <- match("image", sheets_lower)
  if (is.na(idx)) return(character(0))
  img_tbl <- read_excel(path, sheet = sheets_raw[idx])
  if (!all(c("name","media_filename") %in% names(img_tbl))) {
    warnf("Image sheet present but missing required columns 'name' and/or 'media_filename'.")
    return(character(0))
  }
  files <- file.path(media_dir, img_tbl$media_filename)
  exist_map <- stats::setNames(file.exists(files), files)
  valid <- unname(exist_map[files])
  setNames(files[valid], img_tbl$name[valid])
}

# ---------------- Validation ----------------
validate_xlsform <- function(survey, choices, settings, selected_language, ext_choices) {
  # survey columns
  req_cols <- c("type", "name")
  if (!all(req_cols %in% names(survey))) {
    errf("Incomplete XLSForm: 'survey' worksheet must contain columns: %s.", paste(req_cols, collapse = ", "))
  }
  # settings title
  if (!("form_title" %in% names(settings)) || is.na(settings[["form_title"]][1])) {
    errf("Incomplete XLSForm: 'settings$form_title' is missing or NA.")
  }
  # language-specific title warning
  if (!identical(selected_language, "en")) {
    tl <- paste0("form_title::", selected_language)
    if (!(tl %in% names(settings)) || is.na(settings[[tl]][1])) {
      warnf("Requested language '%s' has no 'settings$%s'; falling back to 'form_title'.",
            selected_language, tl)
    }
  }

  # Choice list validation
  # Skip *_from_file types entirely (per requirement)
  types_split <- strsplit(survey$type, "\\s+")
  base_types  <- vapply(types_split, `[`, character(1), 1)
  list_names  <- vapply(types_split, function(x) if (length(x) >= 2) x[2] else NA_character_, NA_character_)

  # Only validate these bases (EXACT match); this excludes *_from_file
  bases_to_check <- c("select_one", "select_multiple", "rank",
                      "select_one_external", "select_multiple_external")

  idx <- which(base_types %in% bases_to_check)
  if (length(idx)) {
    lists <- unique(list_names[idx])
    lists <- lists[!is.na(lists)]
    if (length(lists)) {
      present <- unique(choices$list_name)
      missing <- setdiff(lists, present)
      if (!is.null(ext_choices) && "list_name" %in% names(ext_choices)) {
        present_ext <- unique(ext_choices$list_name)
        missing <- setdiff(missing, present_ext)
      }
      if (length(missing)) {
        errf("Named answer choices not found for lists: %s (check 'choices' or 'external_choices').",
             paste(missing, collapse = ", "))
      }
    }
  }
}

# ---------------- Paragraph helper ----------------
add_par_arial <- function(doc, text, fp_normal) {
  body_add_fpar(
    doc,
    fpar(ftext(text, fp_normal)),
    style = "Normal"
  )
}

# ---------------- Safer multilingual getters ----------------
get_lang_val <- function(row, col, lang) {
  col_lang <- paste0(col, "::", lang)
  if (col_lang %in% names(row)) {
    val <- row[[col_lang]]
  } else if (col %in% names(row)) {
    val <- row[[col]]
  } else {
    val <- NA_character_
  }
  if (length(val) == 0) NA_character_ else val
}

# ---------------- Parse survey rows ----------------
parse_survey_rows <- function(survey, selected_language) {
  questions <- list()
  for (i in seq_len(nrow(survey))) {
    row   <- survey[i, ]
    parts <- strsplit(row$type, "\\s+")[[1]]
    base  <- parts[1]
    if (base == "calculate") next

    attachment <- if (base == "select_one_from_file") parts[2] else NA_character_
    list_name  <- if (length(parts) >= 2 && base %in% c("select_one","select_multiple","rank",
                                                        "select_one_from_file","select_one_external",
                                                        "select_multiple_external")) parts[2] else NA_character_

    q <- list(
      type               = base,
      attachment         = attachment,
      list_name          = list_name,
      name               = row$name,
      label              = get_lang_val(row, "label", selected_language),
      hint               = get_lang_val(row, "hint", selected_language),
      instruction        = if ("instructions" %in% names(row)) get_lang_val(row, "instructions", selected_language) else NA_character_,
      required           = tolower(as.character(row$required)) %in% c("true","yes","1"),
      relevant           = if ("relevant" %in% names(row)) row$relevant else NA,
      constraint_message = get_lang_val(row, "constraint_message", selected_language),
      parameters         = if ("parameters" %in% names(row)) row$parameters else NA_character_,
      repeat_count       = if ("repeat_count" %in% names(row)) row$repeat_count else NA,
      appearance         = if ("appearance" %in% names(row)) row$appearance else NA_character_
    )
    questions[[length(questions) + 1]] <- q
  }
  questions
}

# ---------------- Build choices map ----------------
build_choices_map <- function(choices, selected_language, media_dir, choice_code = FALSE) {
  out <- list()
  has_img <- "image" %in% names(choices)
  all_paths <- if (has_img) unique(file.path(media_dir, na.omit(choices$image))) else character(0)
  exists_map <- if (length(all_paths)) stats::setNames(file.exists(all_paths), all_paths) else list()

  for (ln in unique(choices$list_name)) {
    df <- choices[choices$list_name == ln, , drop = FALSE]
    lbl_col <- if (paste0("label::", selected_language) %in% names(df)) paste0("label::", selected_language) else "label"
    items <- lapply(seq_len(nrow(df)), function(i) {
      row <- df[i, ]
      img_file <- if (has_img && !is.na(row[["image"]])) row[["image"]] else NA_character_
      img_path <- if (!is.na(img_file)) file.path(media_dir, img_file) else NA_character_
      if (!is.na(img_file) && length(img_path) && !isTRUE(exists_map[[img_path]])) {
        warnf("Choice image '%s' for list '%s' item '%s' not found in '%s'.",
              img_file, ln, row$name, media_dir)
        img_path <- NA_character_
      }
      label_text <- as.character(row[[lbl_col]])
      if (isTRUE(choice_code) && !is.na(row$name)) {
        label_text <- sprintf("[%s] %s", row$name, label_text)
      }
      list(
        name  = row$name,
        label = label_text,
        image = img_path
      )
    })
    out[[ln]] <- items
  }
  out
}

# ---------------- Choice table ----------------
choice_items_to_flextable <- function(items, symbol = "\u25EF") {
  df <- data.frame(
    Symbol = rep(symbol, length(items)),
    Choice = vapply(items, function(x) x$label, character(1)),
    stringsAsFactors = FALSE
  )
  ft <- regulartable(df)
  ft <- autofit(ft)
  ft <- align(ft, j = 1, align = "center")
  ft <- valign(ft, j = 1, valign = "center")
  ft
}

# ---------------- Text box helper ----------------
add_text_box <- function(lines) {
  df <- data.frame(text = rep("", lines), stringsAsFactors = FALSE)
  ft <- regulartable(df)
  ft <- delete_part(ft, part = "header")
  ft <- width(ft, j = 1, width = 6.5, unit = "in")
  ft <- border_outer(ft, part = "all", border = fp_border(color = "black", width = 1))
  ft
}

# ---------------- Image sizing ----------------
scaled_external_img <- function(path, max_width_in = 6.5, max_height_in = 3.5, dpi = 96) {
  if (!file.exists(path)) return(external_img(src = path, width = 2))
  info <- tryCatch(magick::image_info(magick::image_read(path)), error = function(e) NULL)
  if (!is.null(info)) {
    w_in <- info$width[1] / dpi
    h_in <- info$height[1] / dpi
    scale <- min(max_width_in / w_in, max_height_in / h_in, 1)
    w_in <- w_in * scale
    h_in <- h_in * scale
    external_img(src = path, width = w_in, height = h_in)
  } else {
    external_img(src = path, width = 3)
  }
}

# ---------------- Line 3 content ----------------
format_line3_blocks <- function(q, choices_map, fp_normal, choice_code, ext_choices) {
  `%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || isTRUE(is.na(a))) b else a
  blks <- list()
  add_fpar <- function(txt) fpar(ftext(txt, fp_normal))
  appearance <- tolower(trimws(q$appearance %||% ""))

  if (identical(q$type, "select_one")) {
    blks <- c(blks, list(add_fpar("(Select only one answer)")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    items <- choices_map[[q$list_name]]
    if (length(items)) blks <- c(blks, list(choice_items_to_flextable(items, symbol = "\u25EF")))
    return(blks)
  }

  if (identical(q$type, "select_multiple")) {
    blks <- c(blks, list(add_fpar("(Select all that apply)")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    items <- choices_map[[q$list_name]]
    if (length(items)) blks <- c(blks, list(choice_items_to_flextable(items, symbol = "\u25A2")))
    return(blks)
  }

  if (identical(q$type, "rank")) {
    blks <- c(blks, list(add_fpar("Rank the items in order")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    items <- choices_map[[q$list_name]]
    if (length(items)) blks <- c(blks, list(choice_items_to_flextable(items, symbol = "\u25EF")))
    return(blks)
  }

  if (identical(q$type, "select_one_from_file")) {
    blks <- c(blks, list(add_fpar(paste0("Answer choices are to be populated from this attachment: ", q$attachment))))
    return(blks)
  }

  if (q$type %in% c("select_one_external", "select_multiple_external")) {
    blks <- c(blks, list(add_fpar("Answer choices are large and provided via external choices list (not enumerated here).")))
    return(blks)
  }

  if (identical(q$type, "range")) {
    blks <- c(blks, list(add_fpar("(Select only one answer)")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    params_raw <- strsplit(q$parameters %||% "", ";")[[1]]
    parts <- strsplit(params_raw, "=", fixed = TRUE)
    k <- trimws(vapply(parts, function(x) x[1] %||% "", character(1)))
    v <- trimws(vapply(parts, function(x) x[2] %||% "", character(1)))
    p_list <- stats::setNames(v, k)
    start <- suppressWarnings(as.numeric(p_list[["start"]] %||% "1"))
    end   <- suppressWarnings(as.numeric(p_list[["end"]]   %||% NA))
    step  <- suppressWarnings(as.numeric(p_list[["step"]]  %||% "1"))
    if (is.na(end)) errf("Question '%s' of type 'range' requires a numeric 'end' parameter.", q$name)
    if (is.na(start) || is.na(step) || step == 0) errf("Question '%s' has invalid 'start' or 'step' for 'range'.", q$name)
    vals  <- seq(start, end, by = step)
    items <- lapply(vals, function(v) list(name = v, label = as.character(v), image = NA_character_))
    blks <- c(blks, list(choice_items_to_flextable(items, symbol = "\u25EF")))
    return(blks)
  }

  if (identical(q$type, "text")) {
    lines <- if (grepl("multiline", appearance, fixed = TRUE)) 6 else 4
    blks <- c(blks, list(add_text_box(lines)))
    return(blks)
  }

  if (identical(q$type, "integer")) {
    blks <- c(blks, list(add_fpar("(Only answer in whole numbers/ integers using numeric characters)")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    blks <- c(blks, list(add_text_box(2)))
    return(blks)
  }

  if (identical(q$type, "decimal")) {
    blks <- c(blks, list(add_fpar("(Only use numeric characters. You may include decimal points)")))
    if (!is.na(q$instruction)) blks <- c(blks, list(add_fpar(q$instruction)))
    blks <- c(blks, list(add_text_box(2)))
    return(blks)
  }

  if (q$type %in% c("date", "time", "dateTime")) {
    blks <- c(blks, list(add_text_box(2)))
    return(blks)
  }

  if (identical(q$type, "image"))  { blks <- c(blks, list(add_fpar("(Either take a new image with the device camera, or draw/ annotate on the device screen)"))); return(blks) }
  if (identical(q$type, "barcode")){ blks <- c(blks, list(add_fpar("(Scan the barcode using the device camera or connected barcode scanner)"))); return(blks) }
  if (identical(q$type, "audio"))  { blks <- c(blks, list(add_fpar("(Capture audio with the device)"))); return(blks) }
  if (identical(q$type, "video"))  { blks <- c(blks, list(add_fpar("(Capture video with the device)"))); return(blks) }
  if (identical(q$type, "geopoint")){ blks <- c(blks, list(add_fpar("(Collect one set of coordinates using device GPS)"))); return(blks) }
  if (identical(q$type, "geotrace")){ blks <- c(blks, list(add_fpar("(Record a line of two or more sets of coordinates using device GPS)"))); return(blks) }
  if (identical(q$type, "geoshape")){ blks <- c(blks, list(add_fpar("(Record a polygon of multiple GPS coordinates using device GPS)"))); return(blks) }

  # file/note/no-op default
  blks
}

# ---------------- Generate docx ----------------
generate_docx <- function(questions, choices_map, settings,
                          selected_language, q_images, number_questions,
                          choice_code, media_dir,
                          fp_normal, fp_title, fp_center,
                          ext_choices = NULL) {
  `%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || isTRUE(is.na(a))) b else a

  doc <- read_docx()

  # Title
  title_col <- paste0("form_title::", selected_language)
  title_val <- if (title_col %in% names(settings) && !is.na(settings[[title_col]][1])) {
    settings[[title_col]][1]
  } else {
    settings[["form_title"]][1]
  }
  doc <- body_add_fpar(doc, fpar(ftext(title_val, fp_title), fp_p = fp_center))
  doc <- body_add_fpar(doc, fpar(ftext("", fp_normal)))

  # Language line
  if (selected_language != "en") {
    doc <- body_add_fpar(doc, fpar(ftext(paste0("Language: ", selected_language), fp_normal)))
    doc <- body_add_fpar(doc, fpar(ftext("", fp_normal)))
  }

  # Mandatory note
  doc <- body_add_fpar(doc, fpar(ftext("Mandatory questions are marked with an asterisk (*) symbol.", fp_normal)))
  doc <- body_add_fpar(doc, fpar(ftext("", fp_normal)))

  # Numbering stack — avoid leading zeros
  render_num <- function(stack) {
    s <- stack[stack > 0]
    if (!length(s)) "" else paste(s, collapse = ".")
  }
  number_stack <- if (number_questions) integer(0) else NULL

  # Precompute question image existence
  q_exist <- if (length(q_images)) stats::setNames(file.exists(unname(q_images)), unname(q_images)) else logical(0)

  # Main loop — build blocks per question and add once
  for (q in questions) {
    blocks <- officer::block_list(...)

    # Structural groups
    if (q$type %in% c("begin_group", "begin_repeat")) {
      if (number_questions) number_stack <- c(number_stack, 0L)
      label <- q$label %||% q$name
      blocks <- append(blocks, list(fpar(ftext(sprintf('%s name "%s" begins here.',
                                                       ifelse(q$type == "begin_group","Group","Repeat group"),
                                                       label), fp_normal))))
      if (!is.na(q$relevant)) blocks <- append(blocks, list(fpar(ftext(paste0("Relevant clause: ", q$relevant), fp_normal))))
      if (q$type == "begin_repeat") {
        cnt <- if (!is.na(q$repeat_count)) q$repeat_count else "repeat these questions as many times as required"
        blocks <- append(blocks, list(fpar(ftext(paste0("Repeat count: ", cnt), fp_normal))))
      }
      blocks <- append(blocks, list(fpar(ftext("", fp_normal))))
      doc <- officer::body_add_blocks(doc, blocks, style = "Normal")
      next
    }

    if (q$type %in% c("end_group", "end_repeat")) {
      label <- q$label %||% q$name
      blocks <- append(blocks, list(fpar(ftext(sprintf('%s name "%s" ends here.',
                                                       ifelse(q$type == "end_group","Group","Repeat group"),
                                                       label), fp_normal))))
      if (number_questions && length(number_stack)) number_stack <- number_stack[-length(number_stack)]
      blocks <- append(blocks, list(fpar(ftext("", fp_normal))))
      doc <- officer::body_add_blocks(doc, blocks, style = "Normal")
      next
    }

    # Numbered label
    lbl_prefix <- ""
    if (number_questions) {
      if (!length(number_stack)) number_stack <- c(0L)
      number_stack[length(number_stack)] <- number_stack[length(number_stack)] + 1L
      num <- render_num(number_stack)
      if (nzchar(num)) lbl_prefix <- paste0(num, ". ")
    }
    req_prefix <- if (isTRUE(q$required)) "* " else ""
    base_label <- q$label %||% q$name %||% ""
    lbl_text <- paste0(lbl_prefix, req_prefix, base_label)
    blocks <- append(blocks, list(fpar(ftext(lbl_text, fp_normal))))

    # Hint
    if (!is.na(q$hint)) blocks <- append(blocks, list(fpar(ftext(q$hint, fp_normal))))

    # Question image
    if (q$name %in% names(q_images)) {
      img_path <- q_images[[q$name]]
      if (isTRUE(q_exist[[img_path]])) {
        blocks <- append(blocks, list(scaled_external_img(img_path, max_width_in = 6.0, max_height_in = 3.0)))
      } else {
        blocks <- append(blocks, list(fpar(ftext(sprintf("(image missing: %s)", img_path), fp_normal))))
      }
    }

    # Line 3 content
    l3 <- format_line3_blocks(q, choices_map, fp_normal, choice_code, ext_choices)
    for (blk in l3) blocks <- append(blocks, list(blk))

    # Constraint message
    if (!is.na(q$constraint_message)) blocks <- append(blocks, list(fpar(ftext(paste0("(", q$constraint_message, ")"), fp_normal))))

    # Relevant clause
    if (!is.na(q$relevant)) blocks <- append(blocks, list(fpar(ftext(paste0("Relevant clause: ", q$relevant), fp_normal))))

    # Blank line
    blocks <- append(blocks, list(fpar(ftext("", fp_normal))))

    # Add blocks
    doc <- officer::body_add_blocks(doc, blocks, style = "Normal")
  }

  doc
}

# ---------------- Save document ----------------
save_docx <- function(doc, path) {
  print(doc, target = path)
}

# ---------------- CLI wrapper ----------------
#' Command-line interface for convertxlsform
#' @param args Character vector like commandArgs(trailingOnly = TRUE)
#' @export
cli_convert <- function(args = commandArgs(trailingOnly = TRUE)) {
  get_opt <- function(flag, default = NULL) {
    hit <- grep(paste0("^", flag, "="), args, value = TRUE)
    if (!length(hit)) return(default)
    sub(paste0("^", flag, "="), "", hit[1])
  }
  xls <- args[1]
  if (is.na(xls) || !nzchar(xls)) {
    errf("Usage: convertxlsform <xlsform.xlsx> [--lang=xx] [--out=path] [--no-number] [--choice-code] [--media=dir]")
  }
  lang  <- get_opt("--lang", "en")
  out   <- get_opt("--out", NULL)
  media <- get_opt("--media", NULL)
  number <- !("--no-number" %in% args)
  ccode  <- "--choice-code" %in% args
  convertxlsform(
    xlsform_path = xls,
    selected_language = lang,
    output_docx = out,
    number_questions = number,
    choice_code = ccode,
    media_dir = media
  )
}
