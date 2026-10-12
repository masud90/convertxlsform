# The public entry point: XLSForm in, Word and/or PDF out. ASCII-only source.

#' Convert an XLSForm to a Word or PDF document
#'
#' Reads an XLSForm (as used by Open Data Kit and KoboToolbox) and writes a
#' formatted paper version for interviewing and review. The default output is
#' a Word document next to the input file. A PDF is made by converting the
#' Word document with LibreOffice (or Microsoft Word through the `doconv`
#' package); see [docx_to_pdf()] and [has_libreoffice()].
#'
#' @section Where the file goes:
#' * `output = NULL`: same folder and base name as `xlsform`, with the new
#'   extension.
#' * `output` is an existing folder: `<base name>.<ext>` inside it.
#' * `output` is a file path: used as given. Its extension (`.docx` or
#'   `.pdf`) sets `format` when `format` is `NULL`, and must agree with
#'   `format` otherwise.
#' * Two formats (`format = c("docx", "pdf")`) write two files with the same
#'   base name. With a file path in `output`, its extension is replaced.
#'
#' An existing file is never replaced unless `overwrite = TRUE`.
#'
#' @param xlsform Path to an XLSForm (`.xlsx`, `.xlsm` or `.xls`).
#' @param output Output file or folder, see Details. `NULL` (default) writes
#'   next to `xlsform`.
#' @param format `"docx"`, `"pdf"`, `"both"` or `c("docx", "pdf")`. `NULL`
#'   (default) takes the format from the extension of `output`, and is
#'   `"docx"` when there is no such extension.
#' @param lang Language to use: a code (`"fr"`), a name (`"French"`) or the
#'   full header suffix (`"French (fr)"`). `NULL` (default) uses the
#'   `default_language` of the settings sheet, or the first language found.
#'   See [xlsform_languages()].
#' @param ... Further options, all named. Defaults are in brackets.
#'   * `nested_serial_numbers` (`TRUE`): number questions inside groups as
#'     1.1, 1.2; `FALSE` gives flat numbers 1, 2, 3.
#'   * `include_relevants` (`TRUE`): show skip logic ("Ask if: ...").
#'   * `include_constraints` (`TRUE`): show constraints.
#'   * `include_hints` (`TRUE`): show hints and guidance hints.
#'   * `include_choice_name` (`TRUE`): show the choice code after each
#'     choice label.
#'   * `show_variable_names` (`FALSE`): show the variable name of each
#'     question.
#'   * `include_hidden` (`FALSE`): include hidden types such as `calculate`,
#'     `start` and `end`.
#'   * `media_root` (folder of `xlsform`): where images, audio, video and
#'     external choice files are looked up.
#'   * `title` (`NULL`): form title; overrides `form_title` of the settings
#'     sheet.
#'   * `page_size` (`"A4"`): `"A4"` or `"letter"`.
#'   * `orientation` (`"portrait"`): `"portrait"` or `"landscape"`.
#'   * `margins_in` (`1`): page margins in inches, above 0 and at most 2.
#'   * `font_family` (`"Arial"`): font for the document.
#'   * `font_size` (`11`): base font size in points, 6 to 20.
#'   * `template` (`NULL`): a reference `.docx` whose styles are used.
#'   * `max_choices` (`30`): choice lists longer than this get a code box
#'     layout instead of a printed list.
#'   * `strings` (`NULL`): named list overriding connective phrases such as
#'     "Ask if".
#'   * `repeat_copies` (`1`): number of blank copies printed for each
#'     repeat group, 1 to 20.
#'   * `font_family_cs` (`NULL`): font for complex scripts; chosen from the
#'     language when `NULL`.
#'   * `overwrite` (`FALSE`): replace existing output files.
#'   * `verbose` (`TRUE`): print progress messages.
#'
#' @return The path or paths of the files written, as a character vector,
#'   invisibly.
#'
#' @seealso [xlsform_outline()] to inspect a form before converting it,
#'   [docx_to_pdf()], and [cli()] for the terminal.
#'
#' @examples
#' path <- system.file("extdata", "minimal", "minimal.xlsx",
#'                     package = "convertxlsform")
#' out <- file.path(tempdir(), "minimal.docx")
#' convertxlsform(path, output = out, verbose = FALSE, overwrite = TRUE)
#'
#' \donttest{
#' # A multilingual form: another language, flat question numbers, US letter
#' path_ml <- system.file("extdata", "multilang", "multilang.xlsx",
#'                        package = "convertxlsform")
#' convertxlsform(path_ml, output = tempdir(), lang = "bn",
#'                nested_serial_numbers = FALSE, page_size = "letter",
#'                verbose = FALSE, overwrite = TRUE)
#'
#' # PDF needs LibreOffice
#' if (interactive() && has_libreoffice()) {
#'   convertxlsform(path, output = tempdir(), format = "pdf",
#'                  verbose = FALSE, overwrite = TRUE)
#' }
#' }
#' @export
convertxlsform <- function(xlsform, output = NULL, format = NULL,
                           lang = NULL, ...) {
  call <- rlang::current_call()
  .with_call(call, {
    started <- Sys.time()
    xlsform <- .cxf_check_input(xlsform)

    opts <- do.call(.cxf_options, c(list(lang = lang), list(...)))
    if (is.null(opts$media_root)) {
      opts$media_root <- dirname(xlsform)
    }
    verbose <- opts$verbose

    targets <- .cxf_targets(xlsform, output, format, opts$overwrite)

    form <- .read_xlsform(xlsform, lang = opts$lang,
                          media_root = opts$media_root, verbose = verbose)
    model <- .build_model(form, opts)
    n_questions <- sum(vapply(model$blocks, function(b) {
      identical(b$block, "question")
    }, NA))
    lang_used <- .language_label(form)
    .cxf_inform(verbose, "Language: {lang_used}; {n_questions} question{?s}")

    docx_path <- targets[["docx"]]
    temp_docx <- is.null(docx_path)
    if (temp_docx) {
      docx_path <- tempfile("cxf_", fileext = ".docx")
      on.exit(unlink(docx_path), add = TRUE)
    }
    .cxf_inform(verbose, "Writing the Word document.")
    .render_docx(model, opts, docx_path)

    if (!is.null(targets[["pdf"]])) {
      .cxf_inform(verbose, "Converting to PDF.")
      docx_to_pdf(docx_path, pdf = targets[["pdf"]], overwrite = TRUE,
                  verbose = FALSE)
    }

    written <- unname(unlist(targets[c("docx", "pdf")]))
    written <- normalizePath(written, winslash = "/", mustWork = FALSE)
    secs <- format(round(as.numeric(difftime(Sys.time(), started,
                                             units = "secs")), 1), nsmall = 1)
    .cxf_inform(verbose, c(v = "Wrote {.file {written}} in {secs} s"))
    invisible(written)
  })
}

#' Language shown in the progress messages: "Name (code)"
#' @noRd
.language_label <- function(form) {
  sfx <- form$lang$suffix
  if (is.null(sfx) || is.na(sfx)) {
    return("single language")
  }
  langs <- form$languages
  i <- match(tolower(sfx), tolower(langs$suffix))
  if (is.na(i)) {
    return(sfx)
  }
  nm <- langs$name[i]
  cd <- langs$code[i]
  if (!is.na(nm) && !is.na(cd)) paste0(nm, " (", cd, ")") else
    if (!is.na(nm)) nm else if (!is.na(cd)) cd else sfx
}

#' Check the input path and return it normalized
#' @noRd
.cxf_check_input <- function(xlsform, call = rlang::caller_env()) {
  if (!(is.character(xlsform) && length(xlsform) == 1L && !is.na(xlsform) &&
        nzchar(xlsform))) {
    .cxf_abort("{.arg xlsform} must be a single file path.", call = call)
  }
  if (!file.exists(xlsform) || dir.exists(xlsform)) {
    wd <- normalizePath(getwd(), winslash = "/", mustWork = FALSE)
    full <- path.expand(xlsform)
    if (!grepl("^(/|[A-Za-z]:)", full)) {
      full <- file.path(wd, full)
    }
    full <- normalizePath(full, winslash = "/", mustWork = FALSE)
    .cxf_abort(c(
      "XLSForm file not found.",
      i = "Looked for {.file {full}}.",
      i = "Working directory: {.file {wd}}."
    ), call = call)
  }
  normalizePath(xlsform, winslash = "/", mustWork = TRUE)
}

#' Work out the output path for each format
#'
#' @return A list with elements `docx` and `pdf`; each is a path or NULL.
#'   Errors if a target exists and `overwrite` is FALSE.
#' @noRd
.cxf_targets <- function(xlsform, output, format, overwrite,
                         call = rlang::caller_env()) {
  known <- c("docx", "pdf")
  if (!is.null(output) && !(is.character(output) && length(output) == 1L &&
                            !is.na(output) && nzchar(output))) {
    .cxf_abort("{.arg output} must be NULL or a single path.", call = call)
  }
  if (!is.null(format)) {
    ok <- is.character(format) && length(format) >= 1L && !anyNA(format)
    format <- tolower(format)
    if (ok && identical(format, "both")) {
      format <- known
    }
    if (!ok || !all(format %in% known)) {
      .cxf_abort(c(
        "{.arg format} must be {.val docx}, {.val pdf} or {.val both}.",
        i = "Got {.val {format}}."
      ), call = call)
    }
    format <- known[known %in% format]
  }

  base <- tools::file_path_sans_ext(basename(xlsform))
  out_dir <- dirname(xlsform)
  out_file <- NULL
  if (!is.null(output)) {
    if (dir.exists(output)) {
      out_dir <- output
    } else {
      ext <- tolower(tools::file_ext(output))
      out_file <- output
      if (!(ext %in% known)) {
        .cxf_abort(c(
          "{.arg output} must end in {.val .docx} or {.val .pdf}, or be an existing folder.",
          i = "Got {.file {output}}."
        ), call = call)
      }
      if (is.null(format)) {
        format <- ext
      } else if (length(format) == 1L && !identical(format, ext)) {
        .cxf_abort(c(
          "{.arg output} and {.arg format} disagree.",
          i = "{.file {output}} is a {ext} file but {.arg format} is {.val {format}}."
        ), call = call)
      }
      out_dir <- dirname(output)
      base <- tools::file_path_sans_ext(basename(output))
    }
  }
  if (is.null(format)) {
    format <- "docx"
  }
  if (!dir.exists(out_dir)) {
    .cxf_abort("Output folder {.file {out_dir}} does not exist.", call = call)
  }

  targets <- list(docx = NULL, pdf = NULL)
  for (f in format) {
    targets[[f]] <- if (!is.null(out_file) && length(format) == 1L) {
      out_file
    } else {
      file.path(out_dir, paste0(base, ".", f))
    }
  }
  existing <- Filter(function(p) !is.null(p) && file.exists(p), targets)
  if (length(existing) > 0L && !isTRUE(overwrite)) {
    first <- existing[[1L]]
    .cxf_abort(c(
      "File {.file {first}} already exists.",
      i = "Use {.code overwrite = TRUE} to replace it."
    ), call = call, class = "cxf_exists_error")
  }
  targets
}
