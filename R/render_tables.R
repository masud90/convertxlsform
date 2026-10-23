# Table layouts that span several questions: choice grids (table-list and
# list-nolabel appearances) and rosters for repeat groups. ASCII-only source.

.cxf_rule_border <- function() {
  officer::fp_border(color = "#BFBFBF", width = 0.5)
}

# ---- grids --------------------------------------------------------------------

#' Most choices a grid takes before falling back to separate questions
#' @noRd
.cxf_grid_max_choices <- 8L

#' Key of a question's choices (names and labels), NA when not gridable
#' @noRd
.grid_key <- function(b) {
  if (!identical(b$block, "question") ||
      !(b$kind %in% c("select_one", "select_multiple"))) {
    return(NA_character_)
  }
  a <- b$answer
  ch <- a$choices
  if (!isTRUE(a$source_found) || is.null(ch) || nrow(ch) == 0L ||
      nrow(ch) > .cxf_grid_max_choices || any(ch$write_in) ||
      any(!vapply(ch$image, is.null, NA))) {
    return(NA_character_)
  }
  paste(b$kind, paste(ch$name, collapse = "\r"),
        paste(ch$label, collapse = "\r"), sep = "\n")
}

#' Index of the last block inside the section at `i` (groups have no end
#' block, so the children are the following blocks that are deeper)
#' @noRd
.section_last <- function(blocks, i) {
  d <- blocks[[i]]$depth
  j <- i
  n <- length(blocks)
  while (j < n && !identical(blocks[[j + 1L]]$block, "section_end") &&
         isTRUE(blocks[[j + 1L]]$depth > d)) {
    j <- j + 1L
  }
  j
}

#' Rows of a `table-list` group, or NULL when it cannot be a grid
#' @noRd
.grid_from_section <- function(blocks, i) {
  b <- blocks[[i]]
  if (isTRUE(b$is_repeat) ||
      !isTRUE(.has_appearance(b$appearance %||% NA_character_,
                              "table-list"))) {
    return(NULL)
  }
  last <- .section_last(blocks, i)
  if (last <= i) {
    return(NULL)
  }
  inner <- blocks[(i + 1L):last]
  keys <- vapply(inner, .grid_key, "")
  if (anyNA(keys) || length(unique(keys)) != 1L) {
    return(NULL)
  }
  list(rows = inner, header = inner[[1L]], last = last)
}

#' A `label` question followed by `list-nolabel` questions, or NULL
#' @noRd
.grid_from_label <- function(blocks, i) {
  b <- blocks[[i]]
  if (!isTRUE(.has_appearance(b$appearance %||% NA_character_, "label"))) {
    return(NULL)
  }
  key <- .grid_key(b)
  if (is.na(key)) {
    return(NULL)
  }
  j <- i
  n <- length(blocks)
  while (j < n) {
    nb <- blocks[[j + 1L]]
    if (!identical(nb$block, "question") ||
        !isTRUE(.has_appearance(nb$appearance %||% NA_character_,
                                "list-nolabel")) ||
        !identical(.grid_key(nb), key)) {
      break
    }
    j <- j + 1L
  }
  if (j == i) {
    return(NULL)
  }
  list(rows = blocks[(i + 1L):j], header = b, last = j)
}

#' Question cell chunks: serial, label, required mark
#' @noRd
.question_chunks <- function(ctx, b, size = ctx$size) {
  fp <- .fp(ctx, size = size)
  ch <- c(
    if (!is.na(b$serial)) list(flextable::as_chunk(
      paste0(b$serial, ". "), props = .fp(ctx, size = size, bold = TRUE)
    )),
    .text_chunks(b$label, fp)
  )
  if (isTRUE(b$required)) {
    ch <- c(ch, list(flextable::as_chunk(
      " *", props = .fp(ctx, size = size, bold = TRUE, color = "#C00000")
    )))
  }
  ch
}

#' Logic lines of several questions, each prefixed with its number
#' @noRd
.rows_logic <- function(rows, ctx) {
  unlist(lapply(rows, function(b) {
    lines <- .logic_lines(b, ctx)
    if (length(lines) && !is.na(b$serial)) {
      lines <- paste0("Q", b$serial, " ", lines)
    }
    lines
  }), use.names = FALSE)
}

#' Grid table: one row per question, one column per choice
#' @noRd
.layout_grid <- function(ctx, rows, header, keep_last = FALSE) {
  ch <- header$answer$choices
  nc <- nrow(ch)
  tw <- ctx$tw
  lab_w <- max(tw * 0.3, min(tw * 0.5, tw - 0.55 * nc))
  cw <- (tw - lab_w) / nc
  nr <- length(rows)
  ft <- .new_table(nr + 1L, nc + 1L)
  small <- ctx$size - 1
  show_code <- isTRUE(ctx$opts$include_choice_name)
  for (k in seq_len(nc)) {
    chunks <- .text_chunks(ch$label[k], .fp(ctx, size = small, bold = TRUE))
    if (show_code) {
      chunks <- c(chunks, list(flextable::as_chunk(
        paste0("\n", ch$name[k]),
        props = .fp(ctx, size = ctx$size - 2, color = .cxf_gray)
      )))
    }
    ft <- .set_cell(ft, 1L, k + 1L, chunks)
  }
  sym <- .cxf_symbols()[[if (header$kind == "select_one") "circle" else
    "square"]]
  fp_sym <- .fp(ctx, size = .symbol_size(sym, ctx$size))
  for (r in seq_len(nr)) {
    ft <- .set_cell(ft, r + 1L, 1L, .question_chunks(ctx, rows[[r]]))
    for (k in seq_len(nc)) {
      ft <- .set_cell(ft, r + 1L, k + 1L,
                      list(flextable::as_chunk(sym, props = fp_sym)))
    }
  }
  ft <- .finish_table(ft, ctx, c(lab_w, rep(cw, nc)), keep_last = keep_last,
                      pad = c(2, 2, 3, 3))
  ft <- flextable::align(ft, j = seq_len(nc) + 1L, align = "center",
                         part = "body")
  ft <- flextable::valign(ft, j = seq_len(nc) + 1L, valign = "center",
                          part = "body")
  ft <- flextable::hline(ft, i = seq_len(nr), border = .cxf_rule_border(),
                         part = "body")
  ft
}

#' Render a grid and the logic lines of its rows
#' @noRd
.render_grid <- function(doc, ctx, g) {
  lines <- .rows_logic(g$rows, ctx)
  ft <- .layout_grid(ctx, g$rows, g$header, keep_last = length(lines) > 0L)
  doc <- .emit(doc, ft)
  .add_logic(doc, ctx, lines)
}

# ---- rosters --------------------------------------------------------------------

#' Question kinds that fit in a roster cell
#' @noRd
.cxf_roster_kinds <- c("text", "number", "select_one", "select_multiple",
                       "date", "time", "datetime", "barcode", "range")

#' Can the blocks of a repeat be drawn as a roster?
#'
#' @param inner The blocks between the repeat section and its end.
#' @param max_cols Most questions (columns) allowed.
#' @noRd
.roster_ok <- function(inner, max_cols) {
  n <- length(inner)
  if (n == 0L || n > max_cols) {
    return(FALSE)
  }
  all(vapply(inner, function(b) {
    identical(b$block, "question") && b$kind %in% .cxf_roster_kinds &&
      is.null(b$image)
  }, NA))
}

#' Should this repeat be a roster?
#' @noRd
.use_roster <- function(inner, opts) {
  switch(opts$repeat_layout,
         roster = .roster_ok(inner, 10L),
         auto = .roster_ok(inner, 6L),
         FALSE)
}

#' Short answer format for a roster header ("dd/mm/yyyy", "code", ...)
#' @noRd
.roster_format <- function(b, ctx) {
  a <- b$answer
  st <- ctx$strings
  switch(
    b$kind,
    select_one = , select_multiple = st$code,
    date = , time = , datetime = gsub("^.*\\((.*)\\)$", "\\1", a$pattern),
    number = if (!is.na(a$range_text %||% NA_character_)) a$range_text else
      NA_character_,
    range = if (!is.na(a$range_text %||% NA_character_)) a$range_text else
      NA_character_,
    NA_character_
  )
}

#' Code list lines under a roster: "Q2.3: 1 = Male, 2 = Female"
#' @noRd
.roster_codes <- function(inner, ctx) {
  out <- character(0)
  for (b in inner) {
    if (!(b$kind %in% c("select_one", "select_multiple"))) next
    ch <- b$answer$choices
    if (is.null(ch) || nrow(ch) == 0L) {
      src <- b$answer$source %||% NA_character_
      if (!is.na(src)) {
        out <- c(out, paste0("Q", b$serial, ": ", ctx$strings$choices_from,
                             " ", src))
      }
      next
    }
    lab <- gsub("\n", " / ", ch$label, fixed = TRUE)
    out <- c(out, paste0("Q", b$serial, ": ",
                         paste(ch$name, "=", lab, collapse = "; ")))
  }
  out
}

#' Roster table: a header row of questions and one blank row per member
#' @noRd
.layout_roster <- function(ctx, inner, nrows, keep_last = FALSE) {
  nq <- length(inner)
  tw <- ctx$tw
  first <- 0.35
  cw <- (tw - first) / nq
  ft <- .new_table(nrows + 1L, nq + 1L)
  small <- ctx$size - 2
  ft <- .set_cell(ft, 1L, 1L, list(flextable::as_chunk(
    "#", props = .fp(ctx, size = small, bold = TRUE)
  )))
  for (j in seq_len(nq)) {
    b <- inner[[j]]
    chunks <- .question_chunks(ctx, b, size = small)
    fmt <- .roster_format(b, ctx)
    if (!is.na(fmt)) {
      chunks <- c(chunks, list(flextable::as_chunk(
        paste0("\n(", fmt, ")"),
        props = .fp(ctx, size = small - 1, italic = TRUE, color = .cxf_gray)
      )))
    }
    ft <- .set_cell(ft, 1L, j + 1L, chunks)
  }
  for (r in seq_len(nrows)) {
    ft <- .set_cell(ft, r + 1L, 1L, list(flextable::as_chunk(
      as.character(r), props = .fp(ctx, size = small, color = .cxf_gray)
    )))
  }
  ft <- .finish_table(ft, ctx, c(first, rep(cw, nq)), keep_last = keep_last,
                      pad = c(2, 2, 3, 3))
  ft <- flextable::border(ft, border = .cxf_box_border(), part = "body")
  ft <- flextable::bg(ft, i = 1L, bg = "#F2F2F2", part = "body")
  ft <- flextable::align(ft, j = 1L, align = "center", part = "body")
  if (nrows > 0L) {
    ft <- flextable::height(ft, i = seq_len(nrows) + 1L, height = 0.32,
                            part = "body")
    ft <- flextable::hrule(ft, i = seq_len(nrows) + 1L, rule = "atleast",
                           part = "body")
  }
  ft
}

#' Render a repeat as a roster: table, code lists, logic lines
#' @noRd
.render_roster <- function(doc, ctx, b, inner) {
  nrows <- b$repeat_count_value
  if (is.null(nrows) || is.na(nrows) || nrows < 1 || nrows > 30) {
    nrows <- ctx$opts$roster_rows
  }
  codes <- .roster_codes(inner, ctx)
  lines <- .rows_logic(inner, ctx)
  ft <- .layout_roster(ctx, inner, as.integer(nrows),
                       keep_last = length(codes) + length(lines) > 0L)
  doc <- .emit(doc, ft)
  for (k in seq_along(codes)) {
    doc <- .add_text(doc, ctx, codes[k], "hint",
                     .fp(ctx, size = ctx$size - 1, color = .cxf_gray),
                     keep_with_next = k < length(codes) || length(lines) > 0L,
                     padding.top = if (k == 1L) 4 else NA)
  }
  .add_logic(doc, ctx, lines)
}
