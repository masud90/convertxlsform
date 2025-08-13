#' Convert an XLSForm to a Word document
#'
#' @importFrom readxl    excel_sheets read_excel
#' @importFrom officer   read_docx body_add_par body_add_fpar body_add_img
#' @importFrom officer   fpar ftext fp_text fp_par fp_border
#' @importFrom commonmark markdown_html
#' @importFrom flextable regulartable border_outer delete_part width body_add_flextable
#' @importFrom magick    image_read image_info
#' @importFrom stats setNames
#' @param xlsform_path Path to the .xlsx XLSForm
#' @param selected_language Language code (defaults to "en")
#' @param output_docx Destination .docx file (defaults to basename(xlsx))
#' @param number_questions Logical: number questions? (defaults to TRUE)
#' @return Invisibly returns the path to the generated .docx
#' @examples
#' \dontrun{
#' # Basic usage, writes survey.docx next to survey.xlsx
#' convertxlsform("survey.xlsx")
#'
#' # Turn off numbering
#' convertxlsform("survey.xlsx", number_questions = FALSE)
#' }
#' @export
convertxlsform <- function(xlsform_path,
                           selected_language = "en",
                           output_docx = NULL,
                           number_questions = TRUE) {
  start_time <- Sys.time()

  # Read sheets and media
  sheets       <- excel_sheets(xlsform_path)
  survey_tbl   <- read_survey(xlsform_path, sheets)
  choices_tbl  <- read_choices(xlsform_path, sheets)
  settings_tbl <- read_settings(xlsform_path, sheets)
  q_images      <- read_images(xlsform_path, sheets)

  # Validate inputs
  validate_xlsform(survey_tbl, choices_tbl, settings_tbl, selected_language)

  # Parse survey and build choices map
  questions    <- parse_survey_rows(survey_tbl, selected_language)
  choices_map  <- build_choices_map(choices_tbl, selected_language, dirname(xlsform_path))

  # Generate Word document
  doc <- generate_docx(
    questions, choices_map, settings_tbl,
    selected_language, q_images, number_questions
  )

  # Determine output path
  if (is.null(output_docx)) {
    base <- tools::file_path_sans_ext(basename(xlsform_path))
    output_docx <- file.path(getwd(), paste0(base, ".docx"))
  }

  # Save document
  save_docx(doc, output_docx)

  # Completion message
  duration <- difftime(Sys.time(), start_time, units = "secs")
  message(sprintf("Document generated in %.2f secs: %s", duration, normalizePath(output_docx)))

  invisible(output_docx)
}

# Read sheets
default_sheet_error <- function() {
  stop("Incomplete xlsform. A required worksheet (survey, choices, settings) is not present")
}
read_survey <- function(path, sheets) {
  if (!"survey" %in% sheets) default_sheet_error()
  read_excel(path, sheet = "survey")
}
read_choices <- function(path, sheets) {
  if (!"choices" %in% sheets) default_sheet_error()
  read_excel(path, sheet = "choices")
}
read_settings <- function(path, sheets) {
  if (!"settings" %in% sheets) default_sheet_error()
  read_excel(path, sheet = "settings")
}
read_images <- function(path, sheets) {
  # Reads 'image' sheet for question images
  if (!"image" %in% sheets) return(character(0))
  img_tbl <- read_excel(path, sheet = "image")
  base_dir <- dirname(path)
  files <- file.path(base_dir, img_tbl$media_filename)
  valid <- file.exists(files)
  setNames(files[valid], img_tbl$name[valid])
}

# Validate xlsform structure
validate_xlsform <- function(survey, choices, settings, selected_language) {
  # Required columns and sheets
  req_cols <- c("type", "name")
  if (!all(req_cols %in% names(survey))) {
    stop("Incomplete xlsform. A required column is not present in the 'survey' worksheet.")
  }
  title_col <- paste0("form_title")
  if (!(title_col %in% names(settings)) || is.na(settings[[title_col]][1])) {
    stop("Incomplete xlsform. The form_title is missing in the settings worksheet.")
  }
  # Validate choice lists
  sel_idx <- grep("^(select_one|select_multiple|rank)", survey$type)
  if (length(sel_idx)) {
    lists <- unique(sapply(strsplit(survey$type[sel_idx], "\\s+"), `[`, 2))
    if (!all(lists %in% choices$list_name)) {
      stop("A named answer choice is not available in the choices worksheet.")
    }
  }
}

#helper: add an Arial‐11 paragraph
add_par_arial <- function(doc, text) {
  body_add_fpar(
    doc,
    fpar(ftext(text, fp_text(font.family = "Arial", font.size = 11))),
    style = "Normal"
  )
}

# Parse survey rows into question objects
parse_survey_rows <- function(survey, selected_language) {
  questions <- list()
  for (i in seq_len(nrow(survey))) {
    row   <- survey[i, ]
    parts <- strsplit(row$type, "\\s+")[[1]]
    base  <- parts[1]
    if (base == "calculate") next
    attachment <- if (base == "select_one_from_file") parts[2] else NA_character_
    list_name  <- if (length(parts) >= 2 && base %in% c("select_one", "select_multiple", "rank", "select_one_from_file")) parts[2] else NA_character_
    get_lang <- function(col) {
      col_lang <- paste0(col, "::", selected_language)
      if (col_lang %in% names(row)) row[[col_lang]] else row[[col]]
    }
    instr <- if ("instructions" %in% names(row)) get_lang("instructions") else NA_character_
    q <- list(
      type               = base,
      attachment         = attachment,
      list_name          = list_name,
      name               = row$name,
      label              = get_lang("label"),
      hint               = get_lang("hint"),
      instruction        = instr,
      required           = tolower(as.character(row$required)) %in% c("true", "yes"),
      relevant           = row$relevant,
      constraint_message = get_lang("constraint_message"),
      parameters         = row$parameters,
      repeat_count       = if (!is.null(row$repeat_count)) row$repeat_count else NA
    )
    questions[[length(questions) + 1]] <- q
  }
  questions
}

# Build lookup for choices including images
build_choices_map <- function(choices, selected_language, base_dir) {
  out <- list()
  img_col <- if ("image" %in% names(choices)) "image" else NULL
  for (ln in unique(choices$list_name)) {
    df <- choices[choices$list_name == ln, ]
    lbl_col <- if (paste0("label::", selected_language) %in% names(df)) paste0("label::", selected_language) else "label"
    items <- lapply(seq_len(nrow(df)), function(i) {
      row <- df[i, ]
      img_file <- if (!is.null(img_col) && !is.na(row[[img_col]])) row[[img_col]] else NA_character_
      img_path <- if (!is.na(img_file)) file.path(base_dir, img_file) else NA_character_
      if (!is.na(img_file) && !file.exists(img_path)) {
        warning(sprintf("Choice image '%s' for list '%s' item '%s' not found in project root", img_file, ln, row$name))
        img_path <- NA_character_
      }
      list(
        name  = row$name,
        label = row[[lbl_col]],
        image = if (!is.na(img_path)) img_path else img_file
      )
    })
    out[[ln]] <- items
  }
  out
}

# Helpers for docx formatting
add_text_box <- function(doc, lines) {
  df <- data.frame(text = rep("", lines), stringsAsFactors = FALSE)
  ft <- regulartable(df)
  # delete header
  ft <- delete_part(ft, part = "header")
  # stretch table to full page width (6.5 inches for 1" margins)
  ft <- width(ft, j = 1, width = 6.5, unit = "in")
  ft <- border_outer(ft, part = "all", border = fp_border(color = "black", width = 1))
  doc <- body_add_flextable(doc, ft)
  doc <- add_par_arial(doc, "")
  doc
}
add_choice_item <- function(doc, item, symbol) {
  # symbol, [name] label, optional image
  txt <- paste0(symbol, " [", item$name, "] ", item$label)
  if (!is.na(item$image) && file.exists(item$image)) {
    doc <- add_par_arial(doc, txt)
    doc <- body_add_img(doc, src = item$image, height = 1, style = "Normal")
  } else if (!is.na(item$image)) {
    # image mentioned but missing, append filename
    txt2 <- paste0(txt, " (image: ", item$image, ")")
    doc <- add_par_arial(doc, txt2)
  } else {
    doc <- add_par_arial(doc, txt)
  }
  doc
}
add_bulleted_list <- function(doc, items, symbol) {
  for (itm in items) {
    doc <- add_choice_item(doc, itm, symbol)
  }
  doc
}

# Format line 3 logic
format_line3 <- function(doc, q, choices_map) {
  switch(q$type,
         select_one = {
           doc <- add_par_arial(doc, "(Select only one answer)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_bulleted_list(doc, choices_map[[q$list_name]], "\u25EF")
         },
         select_one_from_file = {
           add_par_arial(doc, paste0("Answer choices are to be populated from this attachment : ", q$attachment))
         },
         select_multiple = {
           doc <- add_par_arial(doc, "(Select all that apply)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_bulleted_list(doc, choices_map[[q$list_name]], "\u25A2")
         },
         rank = {
           doc <- add_par_arial(doc, "(Select only one answer)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_bulleted_list(doc, choices_map[[q$list_name]], "\u25EF")
         },
         range = {
           doc <- add_par_arial(doc, "(Select only one answer)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           params <- strsplit(q$parameters, ";")[[1]]
           p_list <- setNames(sapply(strsplit(params, "="), `[`, 2), sapply(strsplit(params, "="), `[`, 1))
           start <- as.numeric(p_list["start"][1] %||% 1)
           end   <- as.numeric(p_list["end"][1])
           step  <- as.numeric(p_list["step"][1] %||% 1)
           vals  <- seq(start, end, by = step)
           df    <- lapply(vals, function(v) list(name=v,label=v,image=NA_character_))
           add_bulleted_list(doc, df, "\u25EF")
         },
         text = {
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_text_box(doc, 4)
         },
         integer = {
           doc <- add_par_arial(doc, "(Only answer in whole numbers/ integers using numeric characters)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_text_box(doc, 2)
         },
         decimal = {
           doc <- add_par_arial(doc, "(Only use numeric characters. You may include decimal points)")
           if (!is.na(q$instruction)) doc <- add_par_arial(doc, q$instruction)
           add_text_box(doc, 2)
         },
         date = add_text_box(doc, 2),
         time = add_text_box(doc, 2),
         dateTime = add_text_box(doc, 2),
         image    = add_par_arial(doc, "(Either take a new image with the device camera, or draw/ annotate on the device screen)"),
         barcode  = add_par_arial(doc, "(Scan the barcode using the device camera or connected barcode scanner)"),
         audio    = add_par_arial(doc, "(Capture audio with the device)"),
         video    = add_par_arial(doc, "(Capture video with the device)"),
         geopoint = add_par_arial(doc, "(Collect one set of coordinates using device GPS)"),
         geotrace = add_par_arial(doc, "(Record a line of two or more sets of coordinates using device GPS)"),
         geoshape = add_par_arial(doc, "(Record a polygon of multiple GPS coordinates using device GPS)"),
         file     = doc,
         note     = doc,
         doc
  )
}

# Generate Word document with all questions
generate_docx <- function(questions, choices_map, settings,
                          selected_language, q_images, number_questions) {
  doc <- read_docx()
  # Title
  title_col <- paste0("form_title::", selected_language)
  title_val <- if (title_col %in% names(settings)) {
    settings[[title_col]][1]
  } else {
    settings[["form_title"]][1]
  }
  doc <- body_add_fpar(
    doc,
    fpar(
      ftext(title_val,
            fp_text(font.family = "Arial", font.size = 16, bold = TRUE)),
      fp_p = fp_par(text.align = "center")
    )
  )
  doc <- add_par_arial(doc, "")
  # Language line
  if (selected_language != "en") {
    doc <- add_par_arial(doc, paste0("Language: ", selected_language))
    doc <- add_par_arial(doc, "")
  }

  # note on mandatory questions
  doc <- add_par_arial(
    doc,
    "Mandatory questions are marked with an asterisk (*) symbol."
  )
  doc <- add_par_arial(doc, "")

  # Initialize numbering stack
  if (number_questions) {
    number_stack <- c(0)
  } else {
    number_stack <- NULL
  }

  # Loop through questions
  for (q in questions) {
    # Structural group handling
    if (q$type == "begin_group") {
      if (number_questions) number_stack <- c(number_stack, 0)
      label <- q$label %||% q$name
      doc <- add_par_arial(doc, paste0("Group name \"", label, "\" begins here."))
      if (!is.na(q$relevant)) doc <- add_par_arial(doc, paste0("Relevant clause: ", q$relevant))
      next
    }
    if (q$type == "end_group") {
      label <- q$label %||% q$name
      doc <- add_par_arial(doc, paste0("Group name \"", label, "\" ends here."))
      if (number_questions) number_stack <- number_stack[-length(number_stack)]
      next
    }
    if (q$type == "begin_repeat") {
      if (number_questions) number_stack <- c(number_stack, 0)
      label <- q$label %||% q$name
      doc <- add_par_arial(doc, paste0("Repeat group name \"", label, "\" begins here."))
      if (!is.na(q$relevant)) doc <- add_par_arial(doc, paste0("Relevant clause: ", q$relevant))
      cnt <- if (!is.na(q$repeat_count)) q$repeat_count else "repeat these questions as many times as required"
      doc <- add_par_arial(doc, paste0("Repeat count: ", cnt))
      next
    }
    if (q$type == "end_repeat") {
      label <- q$label %||% q$name
      doc <- add_par_arial(doc, paste0("Repeat group name \"", label, "\" ends here."))
      if (number_questions) number_stack <- number_stack[-length(number_stack)]
      next
    }

    # Regular question numbering and label
    if (q$type == "note") {
      lbl_text <- paste0("Note: ", q$label)
    } else if (number_questions) {
      # increment and render nested number + asterisk if required
      number_stack[length(number_stack)] <- number_stack[length(number_stack)] + 1
      num <- paste(number_stack, collapse = ".")
      lbl_text <- paste0(num, ". ", if (q$required) "* " else "", q$label)
    } else {
      lbl_text <- paste0(if (q$required) "* " else "", q$label)
    }
    doc <- add_par_arial(doc, lbl_text)

    # Hint
    if (!is.na(q$hint)) doc <- add_par_arial(doc, q$hint)

    # Question image
    if (q$name %in% names(q_images)) {
      doc <- body_add_img(doc, src = q_images[[q$name]], height = 2, style = "Normal")
    }

    # Line 3 content
    doc <- format_line3(doc, q, choices_map)

    # Constraint message
    if (!is.na(q$constraint_message)) doc <- add_par_arial(doc, paste0("(", q$constraint_message, ")"))

    # Relevant clause
    if (!is.na(q$relevant)) doc <- add_par_arial(doc, paste0("Relevant clause: ", q$relevant))

    # Blank line
    doc <- add_par_arial(doc, "")
  }

  doc
}

# Save document to file
save_docx <- function(doc, path) {
  print(doc, target = path)
}
