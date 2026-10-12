# Reading an XLSForm into a cxf_form. ASCII-only source.

#' Check the input path
#' @noRd
.check_xlsform_path <- function(path) {
  if (!(is.character(path) && length(path) == 1L && !is.na(path) &&
        nzchar(path))) {
    .cxf_abort("{.arg xlsform} must be a single file path.")
  }
  if (!file.exists(path) || dir.exists(path)) {
    .cxf_abort(c(
      "XLSForm file not found.",
      i = "Looked for {.file {path}}."
    ))
  }
  ext <- tolower(tools::file_ext(path))
  if (!(ext %in% c("xlsx", "xlsm", "xls"))) {
    .cxf_abort(c(
      "{.file {path}} is not an Excel file.",
      i = "An XLSForm must be saved as .xlsx or .xls."
    ))
  }
  normalizePath(path, winslash = "/", mustWork = TRUE)
}

#' Run a readxl call, replacing its errors with a clean one
#' @noRd
.with_readxl <- function(expr, path) {
  tryCatch(
    expr,
    error = function(e) {
      .cxf_abort(c(
        "Cannot read {.file {path}} as an Excel workbook.",
        i = "Check that it is a valid .xlsx or .xls file and is not corrupt."
      ))
    }
  )
}

#' Read one sheet as text with normalized headers
#'
#' The sheet is read from cell A1 so the header is Excel row 1 and each data
#' row keeps its Excel row number in `row`. Fully blank rows are dropped.
#' @noRd
.read_sheet <- function(path, sheet) {
  raw <- .with_readxl(
    readxl::read_excel(
      path, sheet = sheet, col_names = FALSE, col_types = "text",
      range = readxl::cell_limits(c(1, 1), c(NA, NA)),
      .name_repair = "minimal"
    ),
    path
  )
  raw <- as.data.frame(raw, stringsAsFactors = FALSE)
  empty <- data.frame(row = integer(0))
  if (nrow(raw) == 0L || ncol(raw) == 0L) {
    return(empty)
  }
  hdr <- .norm_header(unlist(raw[1L, ], use.names = FALSE))
  hdr[.is_blank(hdr)] <- NA_character_
  keep <- !is.na(hdr) & !duplicated(hdr)
  body <- raw[-1L, keep, drop = FALSE]
  names(body) <- hdr[keep]
  if (ncol(body) == 0L || nrow(body) == 0L) {
    out <- body
    out$row <- integer(nrow(body))
    return(out[0, , drop = FALSE])
  }
  body[] <- lapply(body, function(v) {
    v <- trimws(as.character(v))
    v[!nzchar(v)] <- NA_character_
    v
  })
  rown <- .excel_row(seq_len(nrow(body)))
  blank <- rowSums(!is.na(as.matrix(body))) == 0L
  body <- body[!blank, , drop = FALSE]
  body$row <- rown[!blank]
  rownames(body) <- NULL
  body
}

#' Read the survey, choices and settings sheets
#'
#' @return list(sheets, survey, choices, settings); a missing sheet is NULL.
#' @noRd
.read_raw_sheets <- function(path) {
  sheets <- .with_readxl(readxl::excel_sheets(path), path)
  key <- tolower(trimws(sheets))
  get <- function(nm) {
    i <- match(nm, key)
    if (is.na(i)) NULL else .read_sheet(path, sheets[i])
  }
  list(sheets = sheets, survey = get("survey"), choices = get("choices"),
       settings = get("settings"))
}

#' Build the settings list
#' @noRd
.settings_from_raw <- function(settings, suffix, default_suffix) {
  out <- list(form_title = NA_character_, form_id = NA_character_,
              version = NA_character_, default_language = NA_character_)
  if (is.null(settings) || nrow(settings) == 0L) {
    return(out)
  }
  first <- settings[1L, setdiff(names(settings), "row"), drop = FALSE]
  vals <- lapply(first, function(v) {
    v <- as.character(v)
    if (is.na(v) || !nzchar(v)) NA_character_ else v
  })
  sp <- .split_header(names(first))
  plain <- !is.na(names(first)) & is.na(sp$suffix)
  for (nm in names(first)[plain]) {
    out[[nm]] <- vals[[nm]]
  }
  out$form_title <- .resolve_field(first, "form_title", suffix,
                                   default_suffix)[1L]
  out
}

#' Read an XLSForm
#'
#' @param path Path to an .xlsx or .xls file.
#' @param lang Requested language (code, name or full suffix) or NULL.
#' @param media_root Directory for media and external files; default is the
#'   folder of `path`.
#' @param verbose Print a short summary.
#' @return A `cxf_form`, see `dev/model-spec.md` section 1. Beyond the spec it
#'   carries `sheets` (the sheet names found) and `found` (normalized headers
#'   of the survey and choices sheets), which `.validate_form()` uses.
#' @noRd
.read_xlsform <- function(path, lang = NULL, media_root = NULL,
                          verbose = TRUE) {
  path <- .check_xlsform_path(path)
  .cxf_inform(verbose, "Reading {.file {basename(path)}}")
  media_root <- normalizePath(media_root %||% dirname(path),
                              winslash = "/", mustWork = FALSE)
  raw <- .read_raw_sheets(path)

  sv_raw <- raw$survey %||% data.frame(row = integer(0))
  ch_raw <- raw$choices %||% data.frame(row = integer(0))

  # Languages
  langs <- .detect_languages(names(sv_raw), names(ch_raw))
  default_language <- .settings_from_raw(raw$settings, NA_character_,
                                         NA_character_)$default_language
  lang_info <- .resolve_language(langs, lang, default_language)
  sfx <- lang_info$suffix
  dsfx <- {
    d <- .match_language(langs, default_language)
    if (is.na(d) && nrow(langs) > 0L) langs$suffix[1L] else d
  }

  settings <- .settings_from_raw(raw$settings, sfx, dsfx)

  survey <- .build_survey(sv_raw, sfx, dsfx)
  choices <- .build_choices(ch_raw, sfx, dsfx)

  form <- structure(
    list(
      path = path,
      media_root = media_root,
      settings = settings,
      languages = langs,
      lang = lang_info,
      survey = survey,
      choices = choices,
      sheets = raw$sheets,
      found = list(survey = names(sv_raw), choices = names(ch_raw),
                   has_survey = !is.null(raw$survey),
                   has_choices = !is.null(raw$choices))
    ),
    class = "cxf_form"
  )

  .validate_form(form)
  form$survey <- .build_structure(form$survey, nested_serial_numbers = TRUE)

  form
}

#' Build the survey data.frame
#' @noRd
.build_survey <- function(raw, sfx, dsfx) {
  n <- nrow(raw)
  col <- function(nm) {
    if (nm %in% names(raw)) as.character(raw[[nm]]) else rep(NA_character_, n)
  }
  type <- .squish(col("type"))
  pt <- .parse_type(type)
  out <- data.frame(row = as.integer(raw$row), type = type,
                    stringsAsFactors = FALSE)
  out$base_type <- pt$base_type
  out$type_arg <- pt$type_arg
  out$or_other <- pt$or_other
  out$kind <- .classify_types(pt$base_type)
  out$name <- col("name")
  for (f in .lang_fields()) {
    out[[f]] <- .resolve_field(raw, f, sfx, dsfx)
  }
  req <- .parse_required(col("required"))
  out$required <- req$required
  out$required_expr <- req$required_expr
  for (f in c("relevant", "constraint", "choice_filter", "repeat_count",
              "calculation", "default", "appearance", "parameters",
              "read_only")) {
    out[[f]] <- col(f)
  }
  out$params <- .parse_parameters(out$parameters)
  out$depth <- rep(0L, n)
  out$parent <- rep(NA_integer_, n)
  out$serial <- rep(NA_character_, n)
  out
}

#' Build the choices data.frame
#' @noRd
.build_choices <- function(raw, sfx, dsfx) {
  n <- nrow(raw)
  col <- function(nm) {
    if (nm %in% names(raw)) as.character(raw[[nm]]) else rep(NA_character_, n)
  }
  out <- data.frame(list_name = col("list_name"), name = col("name"),
                    stringsAsFactors = FALSE)
  out$label <- .resolve_field(raw, "label", sfx, dsfx)
  out$image <- .resolve_field(raw, "image", sfx, dsfx)
  out$row <- as.integer(raw$row)
  skip <- c("list_name", "name", "label", "image", "row")
  for (nm in setdiff(names(raw), skip)) {
    out[[nm]] <- raw[[nm]]
  }
  out
}
