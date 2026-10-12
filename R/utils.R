# Internal helpers. ASCII-only source: non-ASCII symbols use \u escapes.

#' @importFrom rlang %||%
#' @noRd
NULL

#' Symbols used in the paper form
#'
#' @return A named list of single-character strings: `circle`, `square`,
#'   `rank` and `tick`.
#' @noRd
.cxf_symbols <- function() {
  list(
    circle = "\u25EF",  # white large circle: choose one (U+25EF)
    square = "\u25A2",  # white square with rounded corners: choose all (U+25A2)
    rank = "\u25A1",    # white square: rank box (U+25A1)
    tick = "\u2610"     # ballot box: acknowledge (U+2610)
  )
}

#' Test for blank values
#'
#' @param x A vector.
#' @return Logical vector, TRUE where `x` is NA or an empty or
#'   whitespace-only string.
#' @noRd
.is_blank <- function(x) {
  if (is.null(x)) {
    return(logical(0))
  }
  out <- is.na(x)
  chr <- as.character(x)
  nonna <- !out
  out[nonna] <- !nzchar(trimws(chr[nonna]))
  out
}

#' Trim and collapse whitespace
#'
#' @param x A character vector (or coercible).
#' @return Character vector with leading and trailing whitespace removed and
#'   internal runs of whitespace collapsed to one space. NA stays NA.
#' @noRd
.squish <- function(x) {
  out <- as.character(x)
  keep <- !is.na(out)
  out[keep] <- trimws(gsub("[[:space:]]+", " ", out[keep]))
  out
}

#' Excel row number for a data row
#'
#' The header is Excel row 1, so data row `i` (1-based, after the header) is
#' Excel row `i + 1`.
#'
#' @param i Integer vector of data row indices.
#' @return Integer vector of Excel row numbers.
#' @noRd
.excel_row <- function(i) {
  as.integer(i) + 1L
}
