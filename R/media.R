# Media files: pixel size from file headers and media records.
# ASCII-only source.

# Missing media files are collected here during model building and reported
# once at the end.
.media_env <- new.env(parent = emptyenv())
.media_env$missing <- character(0)

#' Forget the collected missing media names
#' @noRd
.media_reset_missing <- function() {
  .media_env$missing <- character(0)
  invisible(NULL)
}

#' Missing media names collected so far
#' @noRd
.media_missing <- function() {
  .media_env$missing
}

#' Warn once about all missing media files, then reset
#' @noRd
.media_warn_missing <- function() {
  miss <- unique(.media_env$missing)
  .media_reset_missing()
  if (length(miss) > 0L) {
    .cxf_warn(c(
      "{length(miss)} media file{?s} not found under {.arg media_root}.",
      .bullets(miss),
      i = "The questions are still shown, without the missing files."
    ))
  }
  invisible(miss)
}

#' Big-endian unsigned integer from raw bytes
#' @noRd
.be_int <- function(x) {
  sum(as.numeric(as.integer(x)) * 256^(rev(seq_along(x)) - 1))
}

#' Pixel width and height of a PNG or JPEG
#'
#' Reads only the file header with [readBin()]: the IHDR chunk for PNG and the
#' first start-of-frame marker for JPEG.
#'
#' @param path File path.
#' @return Named numeric vector `c(width, height)`; both NA for other formats,
#'   missing or unreadable files.
#' @noRd
.image_size_px <- function(path) {
  na <- c(width = NA_real_, height = NA_real_)
  if (!is.character(path) || length(path) != 1L || is.na(path) ||
      !file.exists(path) || dir.exists(path)) {
    return(na)
  }
  out <- tryCatch({
    con <- file(path, "rb")
    on.exit(close(con), add = TRUE)
    head <- readBin(con, "raw", 24L)
    png_sig <- as.raw(c(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
    if (length(head) >= 24L && identical(head[1:8], png_sig) &&
        identical(head[13:16], charToRaw("IHDR"))) {
      c(width = .be_int(head[17:20]), height = .be_int(head[21:24]))
    } else if (length(head) >= 2L && head[1L] == as.raw(0xFF) &&
               head[2L] == as.raw(0xD8)) {
      .jpeg_size(con)
    } else {
      na
    }
  }, error = function(e) na, warning = function(w) na)
  if (anyNA(out) || any(out <= 0)) na else out
}

#' Find the start-of-frame marker of a JPEG
#'
#' @param con A binary connection positioned anywhere; rewound past the SOI.
#' @noRd
.jpeg_size <- function(con) {
  na <- c(width = NA_real_, height = NA_real_)
  seek(con, 2L, origin = "start")
  # SOF0-SOF15 except DHT (C4), JPG (C8) and DAC (CC)
  sof <- setdiff(0xC0:0xCF, c(0xC4, 0xC8, 0xCC))
  ff <- as.raw(0xFF)
  repeat {
    b <- readBin(con, "raw", 1L)
    if (length(b) == 0L) {
      return(na)
    }
    if (b != ff) {
      next
    }
    m <- readBin(con, "raw", 1L)
    while (length(m) == 1L && m == ff) {
      m <- readBin(con, "raw", 1L)
    }
    if (length(m) == 0L) {
      return(na)
    }
    mi <- as.integer(m)
    # stuffed byte, TEM, RSTn, SOI: no payload
    if (mi == 0L || mi == 0x01L || (mi >= 0xD0L && mi <= 0xD8L)) {
      next
    }
    # EOI or start of scan before any frame header
    if (mi == 0xD9L || mi == 0xDAL) {
      return(na)
    }
    len <- .be_int(readBin(con, "raw", 2L))
    if (mi %in% sof) {
      d <- readBin(con, "raw", 5L)
      if (length(d) < 5L) {
        return(na)
      }
      return(c(width = .be_int(d[4:5]), height = .be_int(d[2:3])))
    }
    if (len < 2) {
      return(na)
    }
    if (len > 2) {
      skipped <- readBin(con, "raw", len - 2)
      if (length(skipped) < len - 2) {
        return(na)
      }
    }
  }
}

#' Media record for an image file named in the form
#'
#' The file is looked up under `form$media_root` as written, then by its
#' basename. Size is the pixel size at 96 dpi, scaled down (never up) to fit
#' both limits. A missing file gives `found = FALSE` and its name is collected
#' for one warning at the end of model building (see `.media_warn_missing()`).
#'
#' @param form A `cxf_form`.
#' @param file File name as written in the form (NA or blank gives NULL).
#' @param max_height_in,max_width_in Limits in inches.
#' @return NULL, or `list(path, file, found, width_in, height_in)`. The sizes
#'   are NA when the format is not PNG or JPEG.
#' @noRd
.media_record <- function(form, file, max_height_in, max_width_in) {
  if (is.null(file) || length(file) != 1L || .is_blank(file)) {
    return(NULL)
  }
  path <- .find_media_file(form, file)
  if (is.na(path)) {
    .media_env$missing <- c(.media_env$missing, file)
    return(list(path = NA_character_, file = file, found = FALSE,
                width_in = NA_real_, height_in = NA_real_))
  }
  px <- .image_size_px(path)
  w <- unname(px["width"]) / 96
  h <- unname(px["height"]) / 96
  if (!anyNA(c(w, h))) {
    s <- min(1, max_height_in / h, max_width_in / w)
    w <- w * s
    h <- h * s
  }
  list(path = path, file = file, found = TRUE, width_in = w, height_in = h)
}
