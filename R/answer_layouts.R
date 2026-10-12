# Answer layouts: what the interviewer sees below each question. One function
# per display kind, each returning a list of flextables and/or fpars.
# ASCII-only source.

.cxf_box_border <- function() {
  officer::fp_border(color = "#7F7F7F", width = 0.75)
}

#' A blank flextable of nr x nc cells without header
#' @noRd
.new_table <- function(nr, nc) {
  df <- as.data.frame(matrix("", nr, nc), stringsAsFactors = FALSE)
  names(df) <- paste0("c", seq_len(nc))
  flextable::delete_part(flextable::flextable(df), part = "header")
}

#' Put chunks into one cell
#' @noRd
.set_cell <- function(ft, i, j, chunks) {
  flextable::compose(
    ft, i = i, j = j,
    value = flextable::as_paragraph(list_values = chunks), part = "body"
  )
}

#' Common finishing: fonts, fixed layout, widths, borders off, pagination
#'
#' @param widths Column widths in inches (scaled so they sum to the text
#'   width when `fill` is TRUE).
#' @param keep_last Whether the last row keeps with the next paragraph (TRUE
#'   when logic lines follow).
#' @noRd
.finish_table <- function(ft, ctx, widths, keep_last = FALSE,
                          pad = c(1, 1, 3, 3)) {
  widths <- pmax(widths, 0.05)  # never a negative or zero column width
  ft <- flextable::border_remove(ft)
  ft <- flextable::font(ft, fontname = ctx$font, part = "all",
                        cs.family = ctx$cs, hansi.family = ctx$font,
                        eastasia.family = ctx$font)
  ft <- flextable::set_table_properties(ft, layout = "fixed", align = "left")
  ft <- flextable::padding(ft, padding.top = pad[1], padding.bottom = pad[2],
                           padding.left = pad[3], padding.right = pad[4],
                           part = "body")
  ft <- flextable::valign(ft, valign = "top", part = "body")
  for (j in seq_along(widths)) {
    ft <- flextable::width(ft, j = j, width = widths[j], unit = "in")
  }
  nr <- flextable::nrow_part(ft, "body")
  if (nr <= 15L) {
    ft <- flextable::paginate(ft, init = TRUE, hdr_ftr = FALSE)
  }
  if (!keep_last) {
    ft <- flextable::keep_with_next(ft, i = nr, value = FALSE, part = "body")
  }
  ft
}

#' Box the cells (i, j)
#' @noRd
.box_cells <- function(ft, i, j) {
  flextable::border(ft, i = i, j = j, border = .cxf_box_border(),
                    part = "body")
}

#' Empty fill-in box, optionally with text beside it
#'
#' @param box_w Width of the box in inches (`NULL` = full text width).
#' @param lines Box height in lines of 0.3 in.
#' @param beside Text shown to the right of the box, or NA.
#' @noRd
.layout_box <- function(ctx, lines = 1L, box_w = NULL, beside = NA_character_,
                        keep_last = FALSE) {
  tw <- ctx$tw
  if (is.null(box_w) || is.na(beside)) {
    if (is.null(box_w)) box_w <- tw
    widths <- c(box_w, if (tw - box_w > 0.01) tw - box_w)
  } else {
    widths <- c(box_w, tw - box_w)
  }
  ft <- .new_table(1L, length(widths))
  if (!is.na(beside) && length(widths) == 2L) {
    ft <- .set_cell(ft, 1L, 2L, list(flextable::as_chunk(
      beside, props = .fp(ctx, size = ctx$size - 1, color = .cxf_gray)
    )))
  }
  ft <- .finish_table(ft, ctx, widths, keep_last = keep_last)
  ft <- .box_cells(ft, 1L, 1L)
  if (length(widths) == 2L) {
    ft <- flextable::valign(ft, j = 2L, valign = "center", part = "body")
    ft <- flextable::padding(ft, j = 2L, padding.left = 8, part = "body")
  }
  ft <- flextable::height(ft, height = 0.3 * lines, part = "body")
  flextable::hrule(ft, rule = "exact", part = "body")
}

# ---- select ------------------------------------------------------------------

#' Chunks of one choice label (markup, write-in line, image)
#' @noRd
.choice_label_chunks <- function(ctx, label, write_in, image) {
  fp <- .fp(ctx)
  ch <- .text_chunks(label, fp)
  if (isTRUE(write_in)) {
    ch <- c(ch, list(flextable::as_chunk("  ________", props = fp)))
  }
  if (!is.null(image)) {
    usable <- isTRUE(image$found) && !anyNA(c(image$width_in, image$height_in))
    if (usable) {
      ch <- c(ch, list(
        flextable::as_chunk("\n", props = fp),
        flextable::as_image(image$path, width = image$width_in,
                            height = image$height_in, unit = "in")
      ))
    } else {
      ch <- c(ch, list(flextable::as_chunk(
        sprintf("  [%s: %s]", ctx$strings$image, image$file),
        props = .fp(ctx, size = ctx$size - 2, italic = TRUE, color = .cxf_gray)
      )))
    }
  }
  ch
}

#' Source note instead of choices ("Choices from villages.csv")
#' @noRd
.layout_source_note <- function(a, ctx, keep_last) {
  txt <- paste0(ctx$strings$choices_from, " ", a$source)
  note <- officer::fpar(
    officer::ftext(txt, .fp(ctx, size = ctx$size - 1, italic = TRUE,
                            color = .cxf_gray)),
    fp_p = officer::fp_par_lite(word_style = "Normal", padding.bottom = 2,
                                keep_with_next = TRUE)
  )
  list(note, .layout_box(ctx, lines = 1L, keep_last = keep_last))
}

#' Choices as symbol, code and label triples across `ncol` groups
#' @noRd
.layout_choices <- function(a, ctx, ncol, keep_last) {
  ch <- a$choices
  n <- nrow(ch)
  show_code <- isTRUE(ctx$opts$include_choice_name)
  per <- if (show_code) 3L else 2L
  nr <- ceiling(n / ncol)
  tw <- ctx$tw
  gw <- tw / ncol
  sym_w <- 0.32
  indent <- 0.22
  code_w <- if (show_code) {
    mx <- max(nchar(ch$name), 1L)
    min(gw * 0.3, max(0.35, 0.07 * mx + 0.12))
  } else {
    0
  }
  ft <- .new_table(nr, per * ncol)
  sym <- .cxf_symbols()[[switch(a$symbol, circle = "circle",
                                square = "square", "rank")]]
  fp_sym <- .fp(ctx, size = .symbol_size(sym, ctx$size))
  fp_code <- .fp(ctx, size = ctx$size - 2, color = .cxf_gray)
  widths <- numeric(0)
  for (g in seq_len(ncol)) {
    gi <- if (g == 1L) indent else 0
    lw <- gw - sym_w - code_w - gi - if (g < ncol) 0.1 else 0
    widths <- c(widths, sym_w + gi, if (show_code) code_w, lw)
  }
  # fill so that widths sum to the text width (gaps between groups)
  widths[length(widths)] <- widths[length(widths)] + (tw - sum(widths))
  for (k in seq_len(n)) {
    g <- (k - 1L) %/% nr + 1L
    r <- (k - 1L) %% nr + 1L
    j0 <- (g - 1L) * per
    ft <- .set_cell(ft, r, j0 + 1L, list(flextable::as_chunk(sym,
                                                              props = fp_sym)))
    if (show_code) {
      ft <- .set_cell(ft, r, j0 + 2L, list(flextable::as_chunk(
        ch$name[k], props = fp_code
      )))
    }
    ft <- .set_cell(ft, r, j0 + per, .choice_label_chunks(
      ctx, ch$label[k], ch$write_in[k], ch$image[[k]]
    ))
  }
  ft <- .finish_table(ft, ctx, widths, keep_last = keep_last)
  for (g in seq_len(ncol)) {
    j0 <- (g - 1L) * per
    ft <- flextable::valign(ft, j = j0 + 1L, valign = "center",
                            part = "body")
    # rows with an image: symbol and code sit beside the label, at the top
    ks <- which((seq_len(n) - 1L) %/% nr + 1L == g &
                  !vapply(ch$image, is.null, logical(1)))
    if (length(ks)) {
      ft <- flextable::valign(ft, i = (ks - 1L) %% nr + 1L,
                              j = j0 + seq_len(per - 1L), valign = "top",
                              part = "body")
    }
    if (g == 1L) {
      ft <- flextable::padding(ft, j = 1L, padding.left = indent * 72,
                               part = "body")
    }
    if (show_code) {
      ft <- flextable::padding(ft, j = j0 + 2L, padding.top = 3,
                               part = "body")
    }
  }
  list(ft)
}

#' Write-in code box plus a compact code table of every choice
#' @noRd
.layout_code_box <- function(a, ctx, keep_last) {
  ch <- a$choices
  n <- nrow(ch)
  tw <- ctx$tw
  # the "Code:" box
  bx <- .new_table(1L, 3L)
  bx <- .set_cell(bx, 1L, 1L, list(flextable::as_chunk(
    paste0(ctx$strings$code, ":"), props = .fp(ctx, bold = TRUE)
  )))
  bx <- .finish_table(bx, ctx, c(0.8, 1.6, max(tw - 2.4, 0.3)),
                      keep_last = TRUE)
  bx <- .box_cells(bx, 1L, 2L)
  bx <- flextable::valign(bx, j = 1L, valign = "center", part = "body")
  bx <- flextable::height(bx, height = 0.3, part = "body")
  bx <- flextable::hrule(bx, rule = "exact", part = "body")

  ncol <- a$ncol
  nr <- ceiling(n / ncol)
  gw <- tw / ncol
  mx <- max(nchar(ch$name), 1L)
  code_w <- min(gw * 0.35, max(0.4, 0.065 * mx + 0.15))
  fsz <- ctx$size - 2
  ft <- .new_table(nr, 2L * ncol)
  lab <- gsub("\n", " ", ch$label, fixed = TRUE)
  # compose every cell with chunks so that the complex-script font
  # (cs.family) is set on each run
  fp_code <- .fp(ctx, size = fsz, bold = TRUE, color = .cxf_gray)
  fp_lab <- .fp(ctx, size = fsz)
  for (k in seq_len(n)) {
    g <- (k - 1L) %/% nr + 1L
    r <- (k - 1L) %% nr + 1L
    ft <- .set_cell(ft, r, 2L * g - 1L, list(flextable::as_chunk(
      ch$name[k], props = fp_code
    )))
    ft <- .set_cell(ft, r, 2L * g, list(flextable::as_chunk(
      lab[k], props = fp_lab
    )))
  }
  widths <- rep(c(code_w, gw - code_w), ncol)
  ft <- .finish_table(ft, ctx, widths, keep_last = keep_last,
                      pad = c(0.5, 0.5, 3, 3))
  list(bx, ft)
}

.layout_select <- function(a, ctx, keep_last) {
  if (!isTRUE(a$source_found) || nrow(a$choices) == 0L) {
    return(.layout_source_note(a, ctx, keep_last))
  }
  switch(
    a$layout,
    code_box = .layout_code_box(a, ctx, keep_last),
    columns = .layout_choices(a, ctx, a$ncol, keep_last),
    .layout_choices(a, ctx, 1L, keep_last)
  )
}

# ---- other kinds ---------------------------------------------------------------

.layout_range <- function(a, ctx, keep_last) {
  v <- a$values
  if (is.null(v)) {
    return(list(.layout_box(ctx, 1L, box_w = 1.75, beside = a$range_text,
                            keep_last = keep_last)))
  }
  n <- length(v)
  tw <- ctx$tw
  cw <- min(0.6, (tw - 0.5) / n)
  ft <- .new_table(2L, n + 1L)
  nsz <- if (n > 12L) ctx$size - 1 else ctx$size
  for (k in seq_len(n)) {
    ft <- .set_cell(ft, 1L, k, list(flextable::as_chunk(
      format(v[k]), props = .fp(ctx, size = nsz)
    )))
    ft <- .set_cell(ft, 2L, k, list(flextable::as_chunk(
      .cxf_symbols()$circle,
      props = .fp(ctx, size = .symbol_size(.cxf_symbols()$circle, ctx$size))
    )))
  }
  ft <- .finish_table(ft, ctx, c(rep(cw, n), tw - n * cw),
                      keep_last = keep_last, pad = c(0, 0, 0, 0))
  ft <- flextable::align(ft, j = seq_len(n), align = "center", part = "body")
  ft <- flextable::valign(ft, valign = "center", part = "body")
  list(ft)
}

.layout_geo <- function(a, ctx, keep_last) {
  if (identical(a$subtype, "geopoint")) {
    f <- a$fields
    tw <- ctx$tw
    ft <- .new_table(length(f), 3L)
    for (r in seq_along(f)) {
      ft <- .set_cell(ft, r, 1L, list(flextable::as_chunk(
        f[r], props = .fp(ctx)
      )))
    }
    ft <- .finish_table(ft, ctx, c(1.4, 2.4, max(tw - 3.8, 0.3)),
                        keep_last = keep_last,
                        pad = c(2, 2, 3, 3))
    ft <- .box_cells(ft, seq_along(f), 2L)
    ft <- flextable::valign(ft, j = 1L, valign = "center", part = "body")
    ft <- flextable::height(ft, height = 0.3, part = "body")
    ft <- flextable::hrule(ft, rule = "exact", part = "body")
    return(list(ft))
  }
  .layout_prompt(a, ctx)
}

#' Prompt plus "Reference no.: ____" (geotrace, geoshape, media)
#' @noRd
.layout_prompt <- function(a, ctx) {
  list(
    officer::fpar(
      officer::ftext(a$prompt, .fp(ctx, size = ctx$size - 1, italic = TRUE,
                                   color = .cxf_gray)),
      fp_p = officer::fp_par_lite(word_style = "Normal", padding.bottom = 2,
                                  padding.left = 18, keep_with_next = TRUE)
    ),
    officer::fpar(
      officer::ftext(paste0(a$reference, ": ________________"), .fp(ctx)),
      fp_p = officer::fp_par_lite(word_style = "Normal", padding.left = 18)
    )
  )
}

.layout_date <- function(a, ctx) {
  list(officer::fpar(
    officer::ftext(a$pattern, .fp(ctx)),
    fp_p = officer::fp_par_lite(word_style = "Normal", padding.left = 18,
                                padding.top = 3)
  ))
}

.layout_acknowledge <- function(a, ctx) {
  list(officer::fpar(
    officer::ftext(.cxf_symbols()$tick,
                   .fp(ctx, size = .symbol_size(.cxf_symbols()$tick, ctx$size))),
    officer::ftext(paste0("  ", a$prompt), .fp(ctx)),
    fp_p = officer::fp_par_lite(word_style = "Normal", padding.left = 18)
  ))
}

.layout_hidden <- function(a, ctx) {
  txt <- ctx$strings$calculated
  if (!is.null(a$calculation) && !is.na(a$calculation)) {
    txt <- paste0(txt, ctx$colon, a$calculation)
  }
  list(officer::fpar(
    officer::ftext(txt, .fp(ctx, size = ctx$size - 1, italic = TRUE,
                            color = .cxf_gray)),
    fp_p = officer::fp_par_lite(word_style = "Normal", padding.left = 18)
  ))
}

#' Answer layout for a question block
#'
#' @param q A question block.
#' @param ctx Rendering context.
#' @param has_logic Whether logic lines follow (the last table row then keeps
#'   with them).
#' @return A list of flextables and fpars.
#' @noRd
.answer_elements <- function(q, ctx, has_logic = FALSE) {
  a <- q$answer
  switch(
    q$kind,
    select_one = , select_multiple = , rank = .layout_select(a, ctx, has_logic),
    number = list(.layout_box(ctx, 1L, box_w = 1.75, beside = a$range_text,
                              keep_last = has_logic)),
    barcode = , text = , unknown = list(.layout_box(
      ctx, a$lines %||% 2L, keep_last = has_logic
    )),
    range = .layout_range(a, ctx, has_logic),
    date = , time = , datetime = .layout_date(a, ctx),
    geo = .layout_geo(a, ctx, has_logic),
    media = .layout_prompt(a, ctx),
    acknowledge = .layout_acknowledge(a, ctx),
    hidden = .layout_hidden(a, ctx),
    list()
  )
}
