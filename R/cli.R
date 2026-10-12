# Terminal entry point. ASCII-only source.

#' Convert an XLSForm from the terminal
#'
#' Runs [convertxlsform()] with options taken from the command line. Use it
#' from a shell without opening R:
#'
#' ```
#' Rscript -e 'convertxlsform::cli()' form.xlsx --format pdf --lang bn
#' ```
#'
#' Put the XLSForm path and the options after the quoted expression. If a
#' script or CI job needs the exit status, use
#' `Rscript -e 'quit(status = convertxlsform::cli())' form.xlsx`. The plain
#' form above also works for normal use.
#'
#' This function is a terminal interface. Its help and version text go to
#' standard output, and usage errors go to standard error, because that
#' output is its purpose. Progress messages are sent with [message()] and
#' `--quiet` suppresses them (the same as `verbose = FALSE` in
#' [convertxlsform()]).
#'
#' @section Usage:
#' `cli(args)`: `form.xlsx [options]`. Options accept `--name value` and
#' `--name=value`.
#' * `-o`, `--output PATH`: output file or folder (default: next to the
#'   form).
#' * `-f`, `--format docx|pdf|both`: output format (default: from the
#'   extension of `--output`, else `docx`).
#' * `-l`, `--lang LANG`: language code, name or header suffix.
#' * `--flat-numbers`: number questions 1, 2, 3 instead of 1.1, 1.2.
#' * `--no-relevants`: hide skip logic.
#' * `--no-constraints`: hide constraints.
#' * `--no-hints`: hide hints.
#' * `--no-choice-codes`: hide choice codes.
#' * `--show-names`: show variable names.
#' * `--include-hidden`: include calculate, start, end and similar rows.
#' * `--media-root DIR`: folder for media and external files (default: the
#'   folder of the form).
#' * `--title TEXT`: form title.
#' * `--page-size A4|letter` (default `A4`).
#' * `--landscape`: landscape orientation.
#' * `--margins IN`: margins in inches (default 1).
#' * `--font NAME` (default `Arial`) and `--font-size PT` (default 11).
#' * `--template FILE`: reference `.docx` for styles.
#' * `--max-choices N`: longest choice list printed in full (default 30).
#' * `--repeat-copies N`: copies of each repeat group (default 1).
#' * `--overwrite`: replace existing output.
#' * `--quiet`: no progress messages.
#' * `-h`, `--help`: show the help and stop.
#' * `--version`: show the package version and stop.
#'
#' @param args Character vector of arguments. Defaults to the trailing
#'   command-line arguments.
#' @return The exit status invisibly: 0 on success, 1 on a usage error or a
#'   failed conversion. The function does not quit R.
#' @seealso [convertxlsform()]
#' @examples
#' path <- system.file("extdata", "minimal", "minimal.xlsx",
#'                     package = "convertxlsform")
#' cli(c(path, "-o", tempdir(), "--overwrite", "--quiet"))
#' @export
cli <- function(args = commandArgs(trailingOnly = TRUE)) {
  parsed <- tryCatch(
    .parse_cli_args(args),
    cxf_cli_error = function(e) e
  )
  if (inherits(parsed, "error")) {
    .cli_err(paste0("Error: ", conditionMessage(parsed)), "", .cli_usage())
    return(invisible(1L))
  }
  if (isTRUE(parsed$help)) {
    writeLines(.cli_usage())
    return(invisible(0L))
  }
  if (isTRUE(parsed$version)) {
    writeLines(paste("convertxlsform",
                     as.character(utils::packageVersion("convertxlsform"))))
    return(invisible(0L))
  }

  res <- tryCatch(
    do.call(convertxlsform, c(
      list(xlsform = parsed$xlsform, output = parsed$output,
           format = parsed$format, lang = parsed$lang),
      parsed$options
    )),
    error = function(e) e
  )
  if (inherits(res, "error")) {
    msg <- paste0("Error: ", conditionMessage(res))
    if (inherits(res, "cxf_exists_error")) {
      msg <- c(msg, "Use --overwrite to replace the existing file.")
    }
    .cli_err(msg)
    return(invisible(1L))
  }
  invisible(0L)
}

#' Write lines to stderr
#' @noRd
.cli_err <- function(...) {
  writeLines(c(...), con = stderr())
}

#' Raise a command-line usage error
#' @noRd
.cli_stop <- function(msg) {
  rlang::abort(msg, class = "cxf_cli_error", call = NULL)
}

#' Help text
#' @noRd
.cli_usage <- function() {
  c(
    "Convert an XLSForm to a Word or PDF document.",
    "",
    "Usage: Rscript -e 'convertxlsform::cli()' FORM.xlsx [options]",
    "",
    "Options (defaults in brackets):",
    "  -o, --output PATH      Output file or folder [next to the form]",
    "  -f, --format FMT       docx, pdf or both [from --output, else docx]",
    "  -l, --lang LANG        Language code, name or header suffix [form default]",
    "      --flat-numbers     Number questions 1, 2, 3 instead of 1.1, 1.2",
    "      --no-relevants     Hide skip logic",
    "      --no-constraints   Hide constraints",
    "      --no-hints         Hide hints",
    "      --no-choice-codes  Hide choice codes",
    "      --show-names       Show variable names",
    "      --include-hidden   Include calculate, start, end and similar rows",
    "      --media-root DIR   Folder for media and external files [form folder]",
    "      --title TEXT       Form title [form_title of the settings sheet]",
    "      --page-size SIZE   A4 or letter [A4]",
    "      --landscape        Landscape orientation [portrait]",
    "      --margins IN       Margins in inches [1]",
    "      --font NAME        Font family [Arial]",
    "      --font-size PT     Base font size in points [11]",
    "      --template FILE    Reference .docx whose styles are used",
    "      --max-choices N    Longest choice list printed in full [30]",
    "      --repeat-copies N  Copies of each repeat group [1]",
    "      --overwrite        Replace existing output files",
    "      --quiet            No progress messages",
    "  -h, --help             Show this help",
    "      --version          Show the package version",
    "",
    "Options take a value as '--opt value' or '--opt=value'.",
    "For the exit status use: Rscript -e 'quit(status = convertxlsform::cli())' FORM.xlsx"
  )
}

#' Parse command-line arguments
#'
#' @param args Character vector.
#' @return A list: `help`, `version`, `xlsform`, `output`, `format`, `lang`,
#'   and `options` (named list of the remaining `convertxlsform()` options).
#'   Raises an error of class `cxf_cli_error` on bad input.
#' @noRd
.parse_cli_args <- function(args) {
  args <- as.character(args)
  out <- list(help = FALSE, version = FALSE, xlsform = NULL, output = NULL,
              format = NULL, lang = NULL, options = list())

  # Options: name -> list(field/option, kind). Kinds: flag_false, flag_true,
  # string, number, int.
  spec <- list(
    "-o" = "output", "--output" = "output",
    "-f" = "format", "--format" = "format",
    "-l" = "lang", "--lang" = "lang",
    "--media-root" = "media_root", "--title" = "title",
    "--page-size" = "page_size", "--margins" = "margins_in",
    "--font" = "font_family", "--font-size" = "font_size",
    "--template" = "template", "--max-choices" = "max_choices",
    "--repeat-copies" = "repeat_copies"
  )
  flags <- list(
    "--flat-numbers" = list("nested_serial_numbers", FALSE),
    "--no-relevants" = list("include_relevants", FALSE),
    "--no-constraints" = list("include_constraints", FALSE),
    "--no-hints" = list("include_hints", FALSE),
    "--no-choice-codes" = list("include_choice_name", FALSE),
    "--show-names" = list("show_variable_names", TRUE),
    "--include-hidden" = list("include_hidden", TRUE),
    "--landscape" = list("orientation", "landscape"),
    "--overwrite" = list("overwrite", TRUE),
    "--quiet" = list("verbose", FALSE)
  )
  numeric_opts <- c("margins_in", "font_size", "max_choices")
  int_opts <- "repeat_copies"
  positional <- character()

  i <- 1L
  n <- length(args)
  while (i <= n) {
    a <- args[i]
    i <- i + 1L
    if (a %in% c("-h", "--help")) {
      out$help <- TRUE
      next
    }
    if (identical(a, "--version")) {
      out$version <- TRUE
      next
    }
    if (!startsWith(a, "-") || identical(a, "-")) {
      positional <- c(positional, a)
      next
    }

    name <- a
    value <- NULL
    if (startsWith(a, "--") && grepl("=", a, fixed = TRUE)) {
      name <- sub("=.*$", "", a)
      value <- sub("^[^=]*=", "", a)
    }

    if (name %in% names(flags)) {
      if (!is.null(value)) {
        .cli_stop(sprintf("Option %s does not take a value.", name))
      }
      f <- flags[[name]]
      out$options[[f[[1L]]]] <- f[[2L]]
      next
    }
    if (!(name %in% names(spec))) {
      .cli_stop(sprintf("Unknown option: %s", a))
    }
    if (is.null(value)) {
      if (i > n) {
        .cli_stop(sprintf("Option %s needs a value.", name))
      }
      value <- args[i]
      i <- i + 1L
    }
    field <- spec[[name]]

    if (field == "format") {
      v <- tolower(value)
      if (!(v %in% c("docx", "pdf", "both"))) {
        .cli_stop(sprintf("--format must be docx, pdf or both, not '%s'.", value))
      }
      out$format <- if (v == "both") c("docx", "pdf") else v
    } else if (field %in% c("output", "lang")) {
      out[[field]] <- value
    } else if (field == "page_size") {
      v <- switch(tolower(value), a4 = "A4", letter = "letter", NA_character_)
      if (is.na(v)) {
        .cli_stop(sprintf("--page-size must be A4 or letter, not '%s'.", value))
      }
      out$options[[field]] <- v
    } else if (field %in% c(numeric_opts, int_opts)) {
      num <- if (tolower(value) == "inf") Inf else suppressWarnings(as.numeric(value))
      if (is.na(num) || (is.infinite(num) && field != "max_choices")) {
        .cli_stop(sprintf("%s needs a number, not '%s'.", name, value))
      }
      out$options[[field]] <- num
    } else {
      out$options[[field]] <- value
    }
  }

  if (out$help || out$version) {
    return(out)
  }
  if (length(positional) == 0L) {
    .cli_stop("Missing the XLSForm path.")
  }
  if (length(positional) > 1L) {
    .cli_stop(sprintf("Expected one XLSForm path, got %d: %s",
                      length(positional), paste(positional, collapse = ", ")))
  }
  out$xlsform <- positional
  out
}
