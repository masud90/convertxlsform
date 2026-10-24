# Direct WordprocessingML for the common answer tables. Building small
# tables as XML strings is many times faster than going through flextable,
# which matters for forms with hundreds of questions. flextable is still used
# for tables with images, grids and rosters. ASCII-only source.

.cxf_w_ns <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#' Escape text for XML
#' @noRd
.xml_esc <- function(x) {
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  gsub("\"", "&quot;", x, fixed = TRUE)
}

#' Inches to twentieths of a point
#' @noRd
.twips <- function(inches) {
  as.integer(round(inches * 1440))
}

#' Run properties with the document fonts
#' @noRd
.wml_rpr <- function(ctx, size = ctx$size, bold = FALSE, italic = FALSE,
                     color = NA_character_) {
  f <- .xml_esc(ctx$font)
  cs <- .xml_esc(ctx$cs)
  hp <- as.integer(round(size * 2))
  paste0(
    "<w:rPr><w:rFonts w:ascii=\"", f, "\" w:hAnsi=\"", f, "\" w:eastAsia=\"",
    f, "\" w:cs=\"", cs, "\"/>",
    if (isTRUE(bold)) "<w:b/><w:bCs/>",
    if (isTRUE(italic)) "<w:i/><w:iCs/>",
    if (!is.na(color)) paste0("<w:color w:val=\"", sub("^#", "", color),
                              "\"/>"),
    "<w:sz w:val=\"", hp, "\"/><w:szCs w:val=\"", hp, "\"/></w:rPr>"
  )
}

#' One run of plain text (no markup)
#' @noRd
.wml_run <- function(ctx, text, size = ctx$size, bold = FALSE,
                     italic = FALSE, color = NA_character_) {
  paste0("<w:r>", .wml_rpr(ctx, size, bold, italic, color),
         "<w:t xml:space=\"preserve\">", .xml_esc(text), "</w:t></w:r>")
}

#' Runs of label text with markdown, HTML stripping and line breaks
#' @noRd
.wml_text <- function(ctx, text, size = ctx$size, bold = FALSE,
                      italic = FALSE, color = NA_character_) {
  sg <- .text_segments(text)
  if (nrow(sg) == 0L) {
    return("")
  }
  paste(vapply(seq_len(nrow(sg)), function(k) {
    if (sg$br[k]) {
      return("<w:r><w:br/></w:r>")
    }
    .wml_run(ctx, sg$text[k], size, bold || sg$bold[k], italic || sg$italic[k],
             color)
  }, ""), collapse = "")
}

#' A table cell
#'
#' @param width Width in inches.
#' @param runs Run XML for the single paragraph of the cell.
#' @param box Draw a border around the cell?
#' @param valign `"top"`, `"center"` or `"bottom"`.
#' @param align Paragraph alignment, `"left"` or `"center"`.
#' @param pad Padding in points: top, bottom, left, right.
#' @param keep Keep the paragraph with the next one?
#' @noRd
.wml_tc <- function(width, runs = "", box = FALSE, valign = "top",
                    align = "left", pad = c(1, 1, 3, 3), keep = FALSE) {
  bd <- if (box) {
    b <- "w:val=\"single\" w:sz=\"6\" w:space=\"0\" w:color=\"7F7F7F\"/>"
    paste0("<w:tcBorders><w:top ", b, "<w:left ", b, "<w:bottom ", b,
           "<w:right ", b, "</w:tcBorders>")
  } else {
    ""
  }
  tw <- function(pt) as.integer(round(pt * 20))
  paste0(
    "<w:tc><w:tcPr><w:tcW w:w=\"", .twips(width), "\" w:type=\"dxa\"/>", bd,
    "<w:tcMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"", tw(pad[3]),
    "\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"",
    tw(pad[4]), "\" w:type=\"dxa\"/></w:tcMar>",
    "<w:vAlign w:val=\"", valign, "\"/></w:tcPr>",
    "<w:p><w:pPr><w:pStyle w:val=\"Normal\"/>",
    if (keep) "<w:keepNext/>",
    "<w:spacing w:before=\"", tw(pad[1]), "\" w:after=\"", tw(pad[2]),
    "\" w:line=\"240\" w:lineRule=\"auto\"/><w:ind w:left=\"0\" w:right=\"0\"/>",
    "<w:jc w:val=\"", align, "\"/></w:pPr>", runs, "</w:p></w:tc>"
  )
}

#' A table row
#' @param cells Cell XML strings.
#' @param height Row height in inches, or NA for automatic.
#' @param rule `"exact"` or `"atLeast"`.
#' @noRd
.wml_tr <- function(cells, height = NA_real_, rule = "exact") {
  h <- if (is.na(height)) "" else
    paste0("<w:trHeight w:val=\"", .twips(height), "\" w:hRule=\"", rule,
           "\"/>")
  paste0("<w:tr><w:trPr><w:cantSplit/>", h, "</w:trPr>",
         paste(cells, collapse = ""), "</w:tr>")
}

#' A table from rows; returns an element for `.add_elements()`
#' @param widths Column widths in inches.
#' @noRd
.wml_tbl <- function(rows, widths) {
  widths <- pmax(widths, 0.05)
  grid <- paste0("<w:gridCol w:w=\"", .twips(widths), "\"/>", collapse = "")
  xml <- paste0(
    "<w:tbl xmlns:w=\"", .cxf_w_ns, "\"><w:tblPr>",
    "<w:tblW w:w=\"", sum(.twips(widths)), "\" w:type=\"dxa\"/>",
    "<w:jc w:val=\"left\"/><w:tblLayout w:type=\"fixed\"/>",
    "<w:tblCellMar><w:left w:w=\"0\" w:type=\"dxa\"/>",
    "<w:right w:w=\"0\" w:type=\"dxa\"/></w:tblCellMar>",
    "<w:tblLook w:val=\"0000\" w:firstRow=\"0\" w:lastRow=\"0\"",
    " w:firstColumn=\"0\" w:lastColumn=\"0\" w:noHBand=\"1\"",
    " w:noVBand=\"1\"/></w:tblPr><w:tblGrid>", grid, "</w:tblGrid>",
    paste(rows, collapse = ""), "</w:tbl>"
  )
  structure(list(xml = xml), class = "cxf_wml")
}

#' Which rows keep with the next paragraph
#'
#' Short tables (up to 15 rows) are kept on one page; the last row keeps
#' with what follows only when logic lines follow.
#' @noRd
.wml_keep <- function(nr, keep_last) {
  k <- rep(nr <= 15L, nr)
  if (nr > 0L) k[nr] <- isTRUE(keep_last)
  k
}
