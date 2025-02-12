#' Convert ODK XLSForms to Word Documents
#'
#' @description
#' The `convertxlsform` function takes an XLSForm questionnaire from ODK (Open Data Kit)
#' and converts it into a neatly formatted Microsoft Word document.
#' It supports multiple languages (if language-specific columns exist) and applies various formatting
#' options such as bolding required questions, italicizing hints and constraint messages, and more.
#'
#' @param xlsform_path Character. Path to the XLSForm file (in `.xlsx` format).
#' @param selected_language Character. The language to use for label and hint columns. Default is `"en"`.
#' @param output_docx Character. The output Word document filename. If `NULL`, the function uses the XLSForm’s base name appended with `"_PAPER.docx"`.
#'
#' @return Invisibly returns the modified Word document object. The DOCX file is written to disk.
#'
#' @details
#' If the XLSForm does not have language-specific columns (i.e. columns named like `label::<lang>`),
#' the function will use the default columns and will not print the "Selected language" line.
#'
#' @examples
#' \dontrun{
#'   convertxlsform("path/to/your/form.xlsx", selected_language = "English (en)")
#' }
#'
#' @export
convertxlsform <- function(xlsform_path, selected_language = "en", output_docx = NULL) {

  # 1. Set output file name: use the XLSForm base name and append "_PAPER"
  if (is.null(output_docx)) {
    base_name <- tools::file_path_sans_ext(basename(xlsform_path))
    output_docx <- paste0(base_name, "_PAPER.docx")
  }

  # 2. Read the necessary sheets.
  settings <- readxl::read_excel(xlsform_path, sheet = "settings")
  survey   <- readxl::read_excel(xlsform_path, sheet = "survey")
  choices  <- readxl::read_excel(xlsform_path, sheet = "choices")

  # 3. Get the form title.
  form_title <- if ("form_title" %in% names(settings)) {
    settings$form_title[1]
  } else if ("title" %in% names(settings)) {
    settings$title[1]
  } else {
    "Untitled Form"
  }

  # 4. Determine whether language-specific columns exist.
  lang_specified <- any(grepl("^label::", names(survey))) ||
    any(grepl("^hint::", names(survey))) ||
    any(grepl("^constraint::", names(survey)))

  if (!lang_specified) {
    selected_language <- "language not specified"
    label_col_global <- "label"
    hint_col_global <- "hint"
  } else {
    label_col_global <- paste0("label::", selected_language)
    hint_col_global <- paste0("hint::", selected_language)
  }

  # 5. Create a new Word document.
  doc <- officer::read_docx()

  # 6. Define an internal helper function to add formatted paragraphs.
  add_fpar <- function(doc, text, style = "Normal", bold = FALSE, font_size = 11,
                       align = "left", italic = FALSE) {
    fp <- officer::fp_text(font.family = "Arial", bold = bold, font.size = font_size, italic = italic)
    ppr <- officer::fp_par(text.align = align)
    ft <- officer::ftext(text, fp)
    par <- officer::fpar(ft, fp_p = ppr)
    doc <- officer::body_add_fpar(doc, value = par, style = style)
    return(doc)
  }

  # 7. Add the survey title.
  doc <- add_fpar(doc, form_title, style = "graphic title", bold = TRUE, font_size = 16)

  # 8. Add the language line only if language-specific columns exist.
  if (lang_specified) {
    doc <- add_fpar(doc, paste0("Selected language: ", selected_language),
                    style = "graphic title", bold = FALSE, font_size = 10)
    doc <- add_fpar(doc, "", style = "Normal")
  }

  # 8a. Add an extra blank line before the first question.
  doc <- add_fpar(doc, "", style = "Normal", bold = FALSE)

  # 9. Loop over each question.
  for (i in 1:nrow(survey)) {
    q <- survey[i, ]

    # Extract key fields.
    qname      <- if ("name" %in% names(q)) q$name else ""
    full_qtype <- if ("type" %in% names(q)) q$type else ""
    qtype <- if (nchar(full_qtype) > 0) strsplit(full_qtype, " ")[[1]][1] else ""

    # Skip system-generated questions.
    if (tolower(qtype) %in% c("start", "end", "audit", "today", "username", "deviceid")) next

    # For select questions, extract the list name.
    list_name <- if (grepl(" ", full_qtype)) trimws(sub("^[^ ]+\\s+", "", full_qtype)) else ""

    qrelevant    <- if ("relevant" %in% names(q)) q$relevant else ""
    qcalculation <- if ("calculation" %in% names(q)) q$calculation else ""

    # Get language-specific label and hint.
    label_col <- if (lang_specified) label_col_global else "label"
    hint_col  <- if (lang_specified) hint_col_global else "hint"

    qlabel <- if (label_col %in% names(q)) q[[label_col]] else if ("label" %in% names(q)) q$label else ""
    qhint  <- if (hint_col  %in% names(q)) q[[hint_col]]  else if ("hint"  %in% names(q)) q$hint  else ""
    qconstraint_message <- if ("constraint_message" %in% names(q)) q$constraint_message else ""

    lower_full_qtype <- tolower(full_qtype)

    # Prefix required questions with an asterisk.
    required_val <- if ("required" %in% names(q)) q$required else NA
    if (!is.na(required_val) && tolower(as.character(required_val)) == "true") {
      qname <- paste0("*", qname)
    }

    # Handle group/repeat rows.
    if (grepl("^begin_group", lower_full_qtype)) {
      line1 <- paste0("---- Begin group: ", qname, " ----")
      rel_text  <- if (!is.na(qrelevant) && trimws(qrelevant) != "") paste0("Relevance: ", qrelevant) else ""
      calc_text <- if (!is.na(qcalculation) && trimws(qcalculation) != "") paste0("calculation: ", qcalculation) else ""
      line2 <- if (rel_text != "" & calc_text != "") {
        paste0(rel_text, ", ", calc_text)
      } else if (rel_text != "") {
        rel_text
      } else if (calc_text != "") {
        calc_text
      } else {
        ""
      }
      doc <- add_fpar(doc, line1, style = "Normal", bold = TRUE)
      if (line2 != "") doc <- add_fpar(doc, line2, style = "Normal", bold = FALSE)
      doc <- add_fpar(doc, "", style = "Normal")
      next
    } else if (grepl("^end_group", lower_full_qtype)) {
      line <- paste0("---- End group: ", qname, " ----")
      doc <- add_fpar(doc, line, style = "Normal", bold = TRUE)
      doc <- add_fpar(doc, "", style = "Normal")
      next
    } else if (grepl("^begin_repeat", lower_full_qtype)) {
      line1 <- paste0("---- Begin_repeat : ", qname, " ----")
      rel_text  <- if (!is.na(qrelevant) && trimws(qrelevant) != "") paste0("Relevance: ", qrelevant) else ""
      calc_text <- if (!is.na(qcalculation) && trimws(qcalculation) != "") paste0("calculation: ", qcalculation) else ""
      line2 <- if (rel_text != "" & calc_text != "") {
        paste0(rel_text, ", ", calc_text)
      } else if (rel_text != "") {
        rel_text
      } else if (calc_text != "") {
        calc_text
      } else {
        ""
      }
      doc <- add_fpar(doc, line1, style = "Normal", bold = TRUE)
      if (line2 != "") doc <- add_fpar(doc, line2, style = "Normal", bold = FALSE)
      doc <- add_fpar(doc, "", style = "Normal")
      next
    } else if (grepl("^end_repeat", lower_full_qtype)) {
      line <- paste0("---- End repeat: ", qname, " ----")
      doc <- add_fpar(doc, line, style = "Normal", bold = TRUE)
      doc <- add_fpar(doc, "", style = "Normal")
      next
    }

    # Handle note questions.
    if (tolower(qtype) == "note") {
      note_line <- "Note"
      if (!is.na(qrelevant) && trimws(qrelevant) != "") {
        note_line <- paste(note_line, qrelevant)
      }
      doc <- add_fpar(doc, note_line, style = "Normal", bold = TRUE)
      if (!is.na(qlabel) && trimws(qlabel) != "") {
        doc <- add_fpar(doc, qlabel, style = "Normal", bold = FALSE)
      }
      doc <- add_fpar(doc, "", style = "Normal")
      next
    }

    # For regular questions, build the first line.
    line1 <- if (!is.na(qrelevant) && trimws(qrelevant) != "") {
      paste(qname, paste0("(", qtype, ")"), qrelevant, sep = " | ")
    } else {
      paste(qname, paste0("(", qtype, ")"))
    }
    doc <- add_fpar(doc, line1, style = "Normal", bold = TRUE)

    if (!is.na(qlabel) && trimws(qlabel) != "") {
      doc <- add_fpar(doc, qlabel, style = "Normal", bold = FALSE)
    }

    if (!is.na(qhint) && trimws(qhint) != "") {
      doc <- add_fpar(doc, qhint, style = "Normal", bold = FALSE, italic = TRUE)
    }

    # Output the answer area.
    if (tolower(qtype) %in% c("text", "integer", "geopoint")) {
      doc <- add_fpar(doc, "[___ ___ ___ ___ ___]", style = "Normal", bold = FALSE)
    } else if (tolower(qtype) == "calculate") {
      doc <- add_fpar(doc, qcalculation, style = "Normal", bold = FALSE)
    } else if (grepl("^select_one", tolower(full_qtype)) ||
               grepl("^select_multiple", tolower(full_qtype))) {
      current_choices <- dplyr::filter(choices, list_name == !!list_name)
      if (nrow(current_choices) > 0) {
        for (j in 1:nrow(current_choices)) {
          choice <- current_choices[j, ]
          if (lang_specified && paste0("label::", selected_language) %in% names(choice)) {
            choice_label <- choice[[paste0("label::", selected_language)]]
          } else {
            choice_label <- choice$label
          }
          choice_line <- paste(choice$name, choice_label, sep = " - ")
          doc <- add_fpar(doc, choice_line, style = "Normal", bold = FALSE)
        }
      }
    } else {
      doc <- add_fpar(doc, "", style = "Normal", bold = FALSE)
    }

    # Add the constraint message after the answer area.
    if (!is.na(qconstraint_message) && trimws(qconstraint_message) != "") {
      doc <- add_fpar(doc, paste0("(", qconstraint_message, ")"), style = "Normal", bold = FALSE, italic = TRUE)
    }

    # Blank line for spacing between questions.
    doc <- add_fpar(doc, "", style = "Normal", bold = FALSE)
  }

  # Save the Word document.
  print(doc, target = output_docx)
  message("Word document created: ", output_docx)
}
