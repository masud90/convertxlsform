# XLSForm type registry, type parsing and parameter parsing.
# ASCII-only source.

#' Registry of XLSForm types
#'
#' @return A data.frame with `base_type`, `kind` and `subtype`.
#' @noRd
.type_registry <- function() {
  rows <- list(
    c("begin_group", "begin_group", NA),
    c("end_group", "end_group", NA),
    c("begin_repeat", "begin_repeat", NA),
    c("end_repeat", "end_repeat", NA),
    c("select_one", "select_one", NA),
    c("select_one_from_file", "select_one", NA),
    c("select_one_external", "select_one", NA),
    c("select_multiple", "select_multiple", NA),
    c("select_multiple_from_file", "select_multiple", NA),
    c("rank", "rank", NA),
    c("text", "text", NA),
    c("integer", "number", "integer"),
    c("decimal", "number", "decimal"),
    c("range", "range", NA),
    c("date", "date", NA),
    c("time", "time", NA),
    c("datetime", "datetime", NA),
    c("geopoint", "geo", "geopoint"),
    c("geotrace", "geo", "geotrace"),
    c("geoshape", "geo", "geoshape"),
    c("image", "media", "image"),
    c("audio", "media", "audio"),
    c("video", "media", "video"),
    c("file", "media", "file"),
    c("background-audio", "media", "background-audio"),
    c("barcode", "barcode", NA),
    c("acknowledge", "acknowledge", NA),
    c("trigger", "acknowledge", NA),
    c("note", "note", NA)
  )
  hidden <- c("calculate", "hidden", "xml-external", "csv-external", "start",
              "end", "today", "deviceid", "subscriberid", "simserial",
              "phonenumber", "username", "email", "audit", "start-geopoint")
  rows <- c(rows, lapply(hidden, function(h) c(h, "hidden", NA)))
  out <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
  names(out) <- c("base_type", "kind", "subtype")
  out
}

#' Map base types to display kinds
#'
#' @param base_type Character vector of normalized base types.
#' @return Character vector of kinds. Unrecognized or missing types give
#'   `"unknown"`.
#' @noRd
.classify_types <- function(base_type) {
  reg <- .type_registry()
  kind <- reg$kind[match(base_type, reg$base_type)]
  kind[is.na(kind)] <- "unknown"
  kind
}

#' Subtype for each base type (NA when the kind has none)
#' @noRd
.type_subtype <- function(base_type) {
  reg <- .type_registry()
  reg$subtype[match(base_type, reg$base_type)]
}

#' Parse the raw `type` column
#'
#' @param type Character vector of raw types.
#' @return A data.frame with `base_type`, `type_arg` and `or_other`.
#' @noRd
.parse_type <- function(type) {
  n <- length(type)
  base_type <- rep(NA_character_, n)
  type_arg <- rep(NA_character_, n)
  or_other <- rep(FALSE, n)
  x <- .squish(type)
  ok <- !is.na(x) & nzchar(x)
  if (any(ok)) {
    y <- x[ok]
    # "or_other" / "or other" at the end
    oo <- grepl("(^|[[:space:]])or[_ ]other$", y, ignore.case = TRUE,
                perl = TRUE)
    y <- trimws(sub("(^|[[:space:]])or[_ ]other$", "", y,
                    ignore.case = TRUE, perl = TRUE))
    low <- tolower(y)
    bt <- character(length(y))
    rest <- character(length(y))

    is_struct <- grepl("^(begin|end)[ _-]?(group|repeat)($|[[:space:]])", low,
                       perl = TRUE)
    is_sel <- !is_struct &
      grepl("^select[ _](one|multiple)([ _]from[ _]file|[ _]external)?($|[[:space:]])",
            low, perl = TRUE)
    other <- !is_struct & !is_sel

    if (any(is_struct)) {
      m <- regmatches(low[is_struct],
                      regexec("^(begin|end)[ _-]?(group|repeat)(.*)$",
                              low[is_struct], perl = TRUE))
      bt[is_struct] <- vapply(m, function(z) paste0(z[2], "_", z[3]), "")
      rest[is_struct] <- vapply(m, function(z) trimws(z[4]), "")
    }
    if (any(is_sel)) {
      m <- regmatches(low[is_sel],
                      regexec("^(select[ _](?:one|multiple)(?:[ _]from[ _]file|[ _]external)?)(.*)$",
                              low[is_sel], perl = TRUE))
      bt[is_sel] <- vapply(m, function(z) gsub(" ", "_", z[2]), "")
      # keep the original case of the argument
      rest[is_sel] <- substring(y[is_sel],
                                vapply(m, function(z) nchar(z[2]), 1L) + 1L)
      rest[is_sel] <- trimws(rest[is_sel])
    }
    if (any(other)) {
      tok <- sub("[[:space:]].*$", "", y[other])
      bt[other] <- tolower(tok)
      rest[other] <- trimws(substring(y[other], nchar(tok) + 1L))
    }
    arg_types <- c("select_one", "select_one_from_file", "select_one_external",
                   "select_multiple", "select_multiple_from_file", "rank")
    arg <- sub("[[:space:]].*$", "", rest)
    arg[!(bt %in% arg_types) | !nzchar(arg)] <- NA_character_

    base_type[ok] <- bt
    type_arg[ok] <- arg
    or_other[ok] <- oo
  }
  data.frame(base_type = base_type, type_arg = type_arg, or_other = or_other,
             stringsAsFactors = FALSE)
}

#' Parse the `parameters` column
#'
#' Separators between pairs are spaces, commas or semicolons. Numeric values
#' are coerced to numbers; keys are lower-cased.
#'
#' @param x Character vector.
#' @return A list the same length as `x`; each element is a named list
#'   (empty when `x` is NA or empty).
#' @noRd
.parse_parameters <- function(x) {
  x <- as.character(x)
  lapply(x, function(s) {
    out <- stats::setNames(list(), character(0))
    if (is.na(s) || !nzchar(trimws(s))) {
      return(out)
    }
    s <- gsub("[[:space:]]*=[[:space:]]*", "=", s)
    toks <- strsplit(s, "[[:space:],;]+", perl = TRUE)[[1]]
    toks <- toks[nzchar(toks) & grepl("=", toks, fixed = TRUE)]
    for (tk in toks) {
      key <- tolower(trimws(sub("=.*$", "", tk)))
      val <- trimws(sub("^[^=]*=", "", tk))
      if (!nzchar(key) || !nzchar(val)) {
        next
      }
      if (grepl("^[-+]?([0-9]+\\.?[0-9]*|\\.[0-9]+)([eE][-+]?[0-9]+)?$", val)) {
        val <- as.numeric(val)
      }
      out[[key]] <- val
    }
    out
  })
}

#' Parse the `required` column
#'
#' @param x Character vector.
#' @return A data.frame with `required` (lgl) and `required_expr` (chr).
#' @noRd
.parse_required <- function(x) {
  x <- as.character(x)
  low <- tolower(trimws(x))
  blank <- is.na(low) | !nzchar(low)
  yes <- !blank & low %in% c("yes", "true", "1", "true()")
  no <- blank | low %in% c("no", "false", "0", "false()")
  expr <- !yes & !no
  data.frame(
    required = yes | expr,
    required_expr = ifelse(expr, trimws(x), NA_character_),
    stringsAsFactors = FALSE
  )
}
