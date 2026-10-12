#' Convert XLSForm to Word Document
#'
#' @description
#' This is the main entry point of the package. It takes in an XLSForm path,
#' user preferences, and produces a formatted Word document. (Phase 1 skeleton)
#'
#' @param xlsform_path Path to the XLSForm (.xlsx file).
#' @param output_docx Optional output path for the .docx file. Defaults to
#'   the same folder with form title as filename (implemented later).
#' @param lang Language code or name. Default = "en". Accepts "English (en)".
#' @param include_relevants Logical. Whether to include skip logics. Default = TRUE.
#' @param include_choice_name Logical. Whether to include choice names. Default = TRUE.
#' @param nested_serial_numbers Logical. Whether to number questions hierarchically. Default = TRUE.
#' @param page_size Page size for Word. Default = "A4".
#' @param margins_in Page margin in inches. Default = 1.
#' @param media_root Folder path to look for media files. Default = current folder.
#' @param verbose Logical. If TRUE, prints progress messages. Default = TRUE.
#'
#' @return Invisibly returns the path of the generated Word document.
#' @export
#'
#' @examples
#' \dontrun{
#' convertxlsform("form.xlsx", output_docx = "form.docx")
#' }
convertxlsform <- function(
    xlsform_path,
    output_docx = NULL,
    lang = "en",
    include_relevants = TRUE,
    include_choice_name = TRUE,
    nested_serial_numbers = TRUE,
    page_size = "A4",
    margins_in = 1,
    media_root = ".",
    verbose = TRUE
) {
  start_time <- Sys.time()

  if (verbose) {
    message("Starting conversion…")
    message("  Input XLSForm: ", xlsform_path)
    message("  Output DOCX  : ", ifelse(is.null(output_docx), "[auto in later phase]", output_docx))
    message("  Language     : ", lang)
  }

  # ---- Phase 2: read & validate ----
  phase2 <- .read_and_validate_xlsform(xlsform_path, lang = lang, verbose = verbose)
  if (verbose) {
    message("Phase 2 complete: loaded ",
            nrow(phase2$survey), " survey rows; ",
            nrow(phase2$choices), " choices rows; ",
            if (!is.null(phase2$settings)) "settings present" else "no settings")
  }

  # ---- Phase 3: structure & numbering ----
  structure_tbl <- .build_structure(
    survey = phase2$survey,
    nested_serial_numbers = nested_serial_numbers,
    verbose = verbose
  )
  if (verbose) {
    message("Phase 3 complete: structure depth 0–", max(structure_tbl$depth %||% 0, na.rm = TRUE),
            "; questions numbered: ",
            if (nested_serial_numbers) "yes" else "no")
  }

  # Return a placeholder output path to keep the interface stable.
  out <- output_docx %||% tempfile(fileext = ".docx")

  end_time <- Sys.time()
  duration <- difftime(end_time, start_time, units = "secs")
  hms <- .format_duration(duration)
  if (verbose) message("Completed in ", hms)

  invisible(out)
}

.format_duration <- function(duration) {
  secs <- as.integer(duration)
  hrs <- secs %/% 3600
  mins <- (secs %% 3600) %/% 60
  secs <- secs %% 60
  sprintf("%d:%02d:%02d", hrs, mins, secs)
}
