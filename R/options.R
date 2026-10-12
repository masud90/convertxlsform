# Option validation for convertxlsform(). Returns a cxf_options list.

#' Validate and complete the conversion options
#'
#' Takes named option values, fills in defaults for the rest, checks every
#' value, and adds the derived page geometry.
#'
#' @param ... Named option values. Unnamed or unknown names are errors.
#' @return A list of class `cxf_options` with every option, plus
#'   `page_width_in`, `page_height_in` and `text_width_in`.
#' @noRd
.cxf_options <- function(...) {
  supplied <- list(...)
  defaults <- list(
    lang = NULL,
    nested_serial_numbers = TRUE,
    include_relevants = TRUE,
    include_constraints = TRUE,
    include_hints = TRUE,
    include_choice_name = TRUE,
    show_variable_names = FALSE,
    include_hidden = FALSE,
    media_root = NULL,
    title = NULL,
    page_size = c("A4", "letter"),
    orientation = c("portrait", "landscape"),
    margins_in = 1,
    font_family = "Arial",
    font_size = 11,
    template = NULL,
    max_choices = Inf,
    overwrite = FALSE,
    verbose = TRUE
  )

  # Unnamed or unknown arguments
  nms <- names(supplied)
  if (length(supplied) > 0 && (is.null(nms) || any(!nzchar(nms)))) {
    .cxf_abort("All options must be named.")
  }
  unknown <- setdiff(nms, names(defaults))
  if (length(unknown) > 0) {
    .cxf_abort(c(
      "Unknown option{?s}: {.val {unknown}}.",
      i = "Valid options are {.val {names(defaults)}}."
    ))
  }

  opts <- utils::modifyList(defaults, supplied, keep.null = TRUE)

  opts$lang <- .check_optional_string(opts$lang, "lang")
  for (flag in c("nested_serial_numbers", "include_relevants",
                 "include_constraints", "include_hints",
                 "include_choice_name", "show_variable_names",
                 "include_hidden", "overwrite", "verbose")) {
    .check_flag(opts[[flag]], flag)
  }
  opts$media_root <- .check_optional_string(opts$media_root, "media_root")
  opts$title <- .check_optional_string(opts$title, "title")
  opts$page_size <- .check_choice(opts$page_size, "page_size",
                                  c("A4", "letter"))
  opts$orientation <- .check_choice(opts$orientation, "orientation",
                                    c("portrait", "landscape"))
  opts$margins_in <- .check_number(opts$margins_in, "margins_in",
                                   lower = 0, upper = 3,
                                   lower_inclusive = FALSE,
                                   upper_inclusive = FALSE)
  opts$font_family <- .check_string(opts$font_family, "font_family")
  opts$font_size <- .check_number(opts$font_size, "font_size",
                                  lower = 6, upper = 20)
  opts$template <- .check_template(opts$template)
  opts$max_choices <- .check_max_choices(opts$max_choices)

  # Portrait: width is the short side. Landscape swaps the two sides.
  short_side <- switch(opts$page_size, A4 = 8.27, letter = 8.5)
  long_side <- switch(opts$page_size, A4 = 11.69, letter = 11)
  if (opts$orientation == "landscape") {
    opts$page_width_in <- long_side
    opts$page_height_in <- short_side
  } else {
    opts$page_width_in <- short_side
    opts$page_height_in <- long_side
  }
  opts$text_width_in <- opts$page_width_in - 2 * opts$margins_in

  structure(opts, class = "cxf_options")
}

#' Check a single logical flag
#' @noRd
.check_flag <- function(x, name) {
  if (!(is.logical(x) && length(x) == 1L && !is.na(x))) {
    .cxf_abort(c(
      "{.arg {name}} must be TRUE or FALSE.",
      i = "Got {.cls {class(x)}} of length {length(x)}."
    ))
  }
  invisible(TRUE)
}

#' Check a single non-blank string, or NULL
#' @noRd
.check_optional_string <- function(x, name) {
  if (is.null(x)) {
    return(NULL)
  }
  .check_string(x, name)
}

#' Check a single non-blank string
#' @noRd
.check_string <- function(x, name) {
  if (!(is.character(x) && length(x) == 1L && !is.na(x) && nzchar(trimws(x)))) {
    .cxf_abort(c(
      "{.arg {name}} must be a single non-empty string.",
      i = "Got {.cls {class(x)}} of length {length(x)}."
    ))
  }
  x
}

#' Match one of a fixed set of choices
#'
#' A vector equal to the full choice set (the default) selects the first.
#' @noRd
.check_choice <- function(x, name, choices) {
  if (!(is.character(x) && length(x) >= 1L && !anyNA(x))) {
    .cxf_abort("{.arg {name}} must be one of {.val {choices}}.")
  }
  if (length(x) > 1L && setequal(x, choices)) {
    return(choices[1L])
  }
  if (length(x) != 1L || !(x %in% choices)) {
    .cxf_abort(c(
      "{.arg {name}} must be one of {.val {choices}}.",
      i = "Got {.val {x}}."
    ))
  }
  x
}

#' Check a single finite number within bounds
#' @noRd
.check_number <- function(x, name, lower, upper,
                          lower_inclusive = TRUE, upper_inclusive = TRUE) {
  ok <- is.numeric(x) && length(x) == 1L && is.finite(x)
  if (ok) {
    ok <- (if (lower_inclusive) x >= lower else x > lower) &&
      (if (upper_inclusive) x <= upper else x < upper)
  }
  if (!ok) {
    lo <- if (lower_inclusive) "greater than or equal to" else "greater than"
    up <- if (upper_inclusive) "less than or equal to" else "less than"
    .cxf_abort(c(
      "{.arg {name}} must be a single number {lo} {lower} and {up} {upper}.",
      i = "Got {.val {x}}."
    ))
  }
  x
}

#' Check the reference template
#' @noRd
.check_template <- function(x) {
  if (is.null(x)) {
    return(NULL)
  }
  x <- .check_string(x, "template")
  if (!grepl("\\.docx$", x, ignore.case = TRUE)) {
    .cxf_abort(c(
      "{.arg template} must be a .docx file.",
      i = "Got {.val {x}}."
    ))
  }
  if (!file.exists(x)) {
    .cxf_abort(c(
      "{.arg template} file not found.",
      i = "Looked for {.file {x}}."
    ))
  }
  x
}

#' Check max_choices: a positive number or Inf
#' @noRd
.check_max_choices <- function(x) {
  ok <- is.numeric(x) && length(x) == 1L && !is.na(x) && x > 0
  if (!ok) {
    .cxf_abort(c(
      "{.arg max_choices} must be a single positive number or {.code Inf}.",
      i = "Got {.val {x}}."
    ))
  }
  x
}
