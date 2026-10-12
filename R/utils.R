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
    # WGL4 glyphs present in Arial and Liberation Sans (U+25EF and U+2610
    # fall back to other fonts in LibreOffice)
    circle = "\u25CB",  # white circle: choose one (U+25CB)
    square = "\u25A1",  # white square: choose all (U+25A1)
    rank = "\u25A1",    # white square: rank box (U+25A1)
    tick = "\u25A1"     # white square: acknowledge (U+25A1)
  )
}

#' Maximum height in inches of a choice image
#' @noRd
.cxf_choice_image_max_in <- 0.6

#' Maximum height in inches of a question image
#' @noRd
.cxf_question_image_max_in <- 2

#' Font size for a symbol run
#'
#' The hollow circle glyph is drawn much smaller than the square at the same
#' point size, so circles get a larger size.
#' @param sym The symbol character.
#' @param size Body font size in points.
#' @noRd
.symbol_size <- function(sym, size) {
  if (identical(sym, .cxf_symbols()$circle)) size + 7 else size + 3
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
