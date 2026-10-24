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
  ft <- flextable::set_table_properties(ft, layout = "fixed", align = "left",
                                        opts_word = list(split = FALSE))
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
  cells <- .wml_tc(widths[1L], .wml_run(ctx, ""), box = TRUE,
                   keep = keep_last)
  if (length(widths) == 2L) {
    runs <- if (is.na(beside)) "" else
      .wml_run(ctx, beside, size = ctx$size - 1, color = .cxf_gray)
    cells <- c(cells, .wml_tc(widths[2L], runs, valign = "center",
                              pad = c(1, 1, 8, 3), keep = keep_last))
  }
  .wml_tbl(.wml_tr(cells, height = 0.3 * lines), widths)
}

# ---- select ------------------------------------------------------------------

#' Chunks of one choice label (markup, write-in line, image)
#' @noRd
.choice_label_chunks <- function(ctx, label, write_in, image,
                                 skip = NA_character_) {
  fp <- .fp(ctx)
  ch <- .text_chunks(label, fp)
  if (isTRUE(write_in)) {
    ch <- c(ch, list(flextable::as_chunk("  ________", props = fp)))
  }
  if (length(skip) == 1L && !is.na(skip)) {
    ch <- c(ch, list(flextable::as_chunk(
      paste0("  ", ctx$sym$arrow, " ", skip),
      props = .fp(ctx, size = ctx$size - 1, bold = TRUE, color = .cxf_logic)
    )))
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

#' Runs of one choice label (markup, write-in line, go-to instruction)
#' @noRd
.choice_label_wml <- function(ctx, label, write_in, skip = NA_character_) {
  runs <- .wml_text(ctx, label)
  if (isTRUE(write_in)) runs <- paste0(runs, .wml_run(ctx, "  ________"))
  if (length(skip) == 1L && !is.na(skip)) {
    runs <- paste0(runs, .wml_run(ctx, paste0("  ", ctx$sym$arrow, " ", skip),
                                  size = ctx$size - 1, bold = TRUE,
                                  color = .cxf_logic))
  }
  runs
}

#' Column widths of a choice table: symbol, code and label per group
#' @noRd
.choice_widths <- function(ch, ctx, ncol, show_code) {
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
  widths <- numeric(0)
  for (g in seq_len(ncol)) {
    gi <- if (g == 1L) indent else 0
    lw <- gw - sym_w - code_w - gi - if (g < ncol) 0.1 else 0
    widths <- c(widths, sym_w + gi, if (show_code) code_w, lw)
  }
  # fill so that widths sum to the text width (gaps between groups)
  widths[length(widths)] <- widths[length(widths)] + (tw - sum(widths))
  list(widths = widths, indent = indent)
}

#' Choices as symbol, code and label triples across `ncol` groups
#'
#' Choice images need flextable; other lists are written as WML directly.
#' @noRd
.layout_choices <- function(a, ctx, ncol, keep_last) {
  ch <- a$choices
  if (any(!vapply(ch$image, is.null, NA))) {
    return(.layout_choices_ft(a, ctx, ncol, keep_last))
  }
  n <- nrow(ch)
  show_code <- isTRUE(ctx$opts$include_choice_name)
  per <- if (show_code) 3L else 2L
  nr <- ceiling(n / ncol)
  cw <- .choice_widths(ch, ctx, ncol, show_code)
  widths <- cw$widths
  sym <- .cxf_symbols()[[switch(a$symbol, circle = "circle",
                                square = "square", "rank")]]
  sym_run <- .wml_run(ctx, sym, size = .symbol_size(sym, ctx$size))
  keep <- .wml_keep(nr, keep_last)
  rows <- character(nr)
  for (r in seq_len(nr)) {
    cells <- character(0)
    for (g in seq_len(ncol)) {
      k <- (g - 1L) * nr + r
      j0 <- (g - 1L) * per
      has <- k <= n
      left <- if (g == 1L) cw$indent * 72 else 3
      cells <- c(cells, .wml_tc(widths[j0 + 1L], if (has) sym_run else "",
                                valign = "center", pad = c(1, 1, left, 3),
                                keep = keep[r]))
      if (show_code) {
        cells <- c(cells, .wml_tc(
          widths[j0 + 2L],
          if (has) .wml_run(ctx, ch$name[k], size = ctx$size - 2,
                            color = .cxf_gray) else "",
          pad = c(3, 1, 3, 3), keep = keep[r]
        ))
      }
      cells <- c(cells, .wml_tc(
        widths[j0 + per],
        if (has) .choice_label_wml(ctx, ch$label[k], ch$write_in[k],
                                   ch$skip[k] %||% NA_character_) else "",
        keep = keep[r]
      ))
    }
    rows[r] <- .wml_tr(cells)
  }
  list(.wml_tbl(rows, widths))
}

#' Choices with images, through flextable
#' @noRd
.layout_choices_ft <- function(a, ctx, ncol, keep_last) {
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
      ctx, ch$label[k], ch$write_in[k], ch$image[[k]],
      skip = ch$skip[k] %||% NA_character_
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
  bw <- c(0.8, 1.6, max(tw - 2.4, 0.3))
  bx <- .wml_tbl(.wml_tr(c(
    .wml_tc(bw[1L], .wml_run(ctx, paste0(ctx$strings$code, ":"), bold = TRUE),
            valign = "center", keep = TRUE),
    .wml_tc(bw[2L], "", box = TRUE, keep = TRUE),
    .wml_tc(bw[3L], "", keep = TRUE)
  ), height = 0.3), bw)

  ncol <- a$ncol
  nr <- ceiling(n / ncol)
  gw <- tw / ncol
  mx <- max(nchar(ch$name), 1L)
  code_w <- min(gw * 0.35, max(0.4, 0.065 * mx + 0.15))
  fsz <- ctx$size - 2
  lab <- gsub("\n", " ", ch$label, fixed = TRUE)
  widths <- rep(c(code_w, gw - code_w), ncol)
  keep <- .wml_keep(nr, keep_last)
  pad <- c(0.5, 0.5, 3, 3)
  rows <- character(nr)
  for (r in seq_len(nr)) {
    cells <- character(0)
    for (g in seq_len(ncol)) {
      k <- (g - 1L) * nr + r
      has <- k <= n
      code <- if (has) .wml_run(ctx, ch$name[k], size = fsz, bold = TRUE,
                                color = .cxf_gray) else ""
      txt <- ""
      if (has) {
        txt <- .wml_run(ctx, lab[k], size = fsz)
        sk <- ch$skip[k] %||% NA_character_
        if (!is.na(sk)) {
          txt <- paste0(txt, .wml_run(ctx, paste0("  ", ctx$sym$arrow, " ", sk),
                                      size = fsz, bold = TRUE,
                                      color = .cxf_logic))
        }
      }
      cells <- c(cells,
                 .wml_tc(widths[2L * g - 1L], code, pad = pad, keep = keep[r]),
                 .wml_tc(widths[2L * g], txt, pad = pad, keep = keep[r]))
    }
    rows[r] <- .wml_tr(cells)
  }
  list(bx, .wml_tbl(rows, widths))
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
  widths <- c(rep(cw, n), tw - n * cw)
  nsz <- if (n > 12L) ctx$size - 1 else ctx$size
  sym <- .cxf_symbols()$circle
  sym_run <- .wml_run(ctx, sym, size = .symbol_size(sym, ctx$size))
  pad <- c(0, 0, 0, 0)
  nums <- c(vapply(seq_len(n), function(k) {
    .wml_tc(cw, .wml_run(ctx, format(v[k]), size = nsz), valign = "center",
            align = "center", pad = pad, keep = TRUE)
  }, ""), .wml_tc(widths[n + 1L], "", pad = pad, keep = TRUE))
  syms <- c(vapply(seq_len(n), function(k) {
    .wml_tc(cw, sym_run, valign = "center", align = "center", pad = pad,
            keep = keep_last)
  }, ""), .wml_tc(widths[n + 1L], "", pad = pad, keep = keep_last))
  list(.wml_tbl(c(.wml_tr(nums), .wml_tr(syms)), widths))
}

.layout_geo <- function(a, ctx, keep_last) {
  if (identical(a$subtype, "geopoint")) {
    f <- a$fields
    tw <- ctx$tw
    widths <- c(1.4, 2.4, max(tw - 3.8, 0.3))
    keep <- .wml_keep(length(f), keep_last)
    pad <- c(2, 2, 3, 3)
    rows <- vapply(seq_along(f), function(r) {
      .wml_tr(c(
        .wml_tc(widths[1L], .wml_run(ctx, f[r]), valign = "center", pad = pad,
                keep = keep[r]),
        .wml_tc(widths[2L], "", box = TRUE, pad = pad, keep = keep[r]),
        .wml_tc(widths[3L], "", pad = pad, keep = keep[r])
      ), height = 0.3)
    }, "")
    return(list(.wml_tbl(rows, widths)))
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
