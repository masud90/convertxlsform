# The Word renderer: turns a cxf_model into a .docx with officer and
# flextable. ASCII-only source.

#' Parse label text into formatted segments
#'
#' Strips HTML tags, turns `**bold**`, `*italic*` and `_italic_` into
#' formatting flags and keeps line breaks as separate segments.
#'
#' @param text A single string (NA gives no segments).
#' @return A data.frame with `text`, `bold`, `italic` and `br` (TRUE for a
#'   line break; the text is then empty).
#' @noRd
.text_segments <- function(text) {
  empty <- data.frame(text = character(0), bold = logical(0),
                      italic = logical(0), br = logical(0),
                      stringsAsFactors = FALSE)
  if (is.null(text) || length(text) != 1L || is.na(text) || !nzchar(text)) {
    return(empty)
  }
  # plain text (the usual case) needs no parsing
  if (!grepl("[*_<>&#\\[\r\n]", text)) {
    return(data.frame(text = as.character(text), bold = FALSE, italic = FALSE,
                      br = FALSE, stringsAsFactors = FALSE))
  }
  x <- .norm_newlines(as.character(text))
  x <- gsub("<[[:space:]]*br[[:space:]]*/?[[:space:]]*>", "\n", x,
            ignore.case = TRUE)
  x <- gsub("</?[A-Za-z][^>]*>", "", x)
  x <- gsub("&nbsp;", " ", x, fixed = TRUE)
  x <- gsub("&lt;", "<", x, fixed = TRUE)
  x <- gsub("&gt;", ">", x, fixed = TRUE)
  x <- gsub("&quot;", "\"", x, fixed = TRUE)
  x <- gsub("&amp;", "&", x, fixed = TRUE)
  x <- gsub("\\[([^]\n]+)\\]\\([^)\n]*\\)", "\\1", x, perl = TRUE)
  pat <- paste0(
    "\\*\\*(?=\\S)(.+?)(?<=\\S)\\*\\*",
    "|\\*(?=[^\\s*])(.+?)(?<=[^\\s*])\\*",
    "|(?<![[:alnum:]_])_(?=[^\\s_])(.+?)(?<=[^\\s_])_(?![[:alnum:]_])"
  )
  lines <- strsplit(x, "\n", fixed = TRUE)[[1L]]
  if (length(lines) == 0L) {
    return(empty)
  }
  segs <- list()
  add <- function(t, b = FALSE, i = FALSE, br = FALSE) {
    segs[[length(segs) + 1L]] <<- data.frame(
      text = t, bold = b, italic = i, br = br, stringsAsFactors = FALSE
    )
  }
  for (li in seq_along(lines)) {
    if (li > 1L) add("", br = TRUE)
    s <- lines[li]
    first <- length(segs) + 1L
    is_head <- grepl("^[[:space:]]*#{1,6}[[:space:]]+", s)
    if (is_head) s <- sub("^[[:space:]]*#{1,6}[[:space:]]+", "", s)
    while (nzchar(s)) {
      m <- regexec(pat, s, perl = TRUE)[[1L]]
      if (m[1L] == -1L) {
        add(s)
        break
      }
      st <- m[1L]
      len <- attr(m, "match.length")[1L]
      if (st > 1L) add(substr(s, 1L, st - 1L))
      grp <- vapply(2:4, function(g) {
        if (m[g] > 0L) substr(s, m[g], m[g] + attr(m, "match.length")[g] - 1L)
        else ""
      }, "")
      g <- which(nzchar(grp))[1L]
      add(grp[g], b = g == 1L, i = g != 1L)
      s <- substr(s, st + len, nchar(s))
    }
    if (is_head) {
      for (k in seq_len(length(segs) - first + 1L) + first - 1L) {
        segs[[k]]$bold <- TRUE
      }
    }
  }
  do.call(rbind, segs)
}

#' Runs for officer::fpar() from label text
#'
#' @param text A single string with optional markdown and `\n` line breaks.
#' @param fp An `officer::fp_text` (or `fp_text_lite`) for plain text.
#' @return A list of runs (possibly empty).
#' @noRd
.text_runs <- function(text, fp) {
  sg <- .text_segments(text)
  lapply(seq_len(nrow(sg)), function(k) {
    if (sg$br[k]) {
      return(officer::run_linebreak())
    }
    p <- fp
    if (sg$bold[k]) p$bold <- TRUE
    if (sg$italic[k]) p$italic <- TRUE
    officer::ftext(sg$text[k], p)
  })
}

#' Chunks for flextable::as_paragraph() from label text
#' @noRd
.text_chunks <- function(text, fp) {
  sg <- .text_segments(text)
  lapply(seq_len(nrow(sg)), function(k) {
    if (sg$br[k]) {
      return(flextable::as_chunk("\n", props = fp))
    }
    p <- fp
    if (sg$bold[k]) p$bold <- TRUE
    if (sg$italic[k]) p$italic <- TRUE
    flextable::as_chunk(sg$text[k], props = p)
  })
}

# ---- context ---------------------------------------------------------------

#' Rendering context shared by the helpers
#' @noRd
.render_ctx <- function(model, opts, doc) {
  si <- officer::styles_info(doc)
  pstyles <- si$style_name[si$style_type == "paragraph"]
  st <- model$meta$strings
  colon <- if (grepl(" : ", st$legend_ask_if, fixed = TRUE)) " : " else ": "
  list(
    opts = opts, meta = model$meta, strings = st, colon = colon,
    font = opts$font_family,
    cs = model$meta$font_family_cs %||% opts$font_family,
    size = opts$font_size, tw = opts$text_width_in,
    pstyles = pstyles, sym = .cxf_symbols()
  )
}

#' Text properties with the document fonts
#' @noRd
.fp <- function(ctx, size = ctx$size, bold = NA, italic = NA, color = NA) {
  officer::fp_text_lite(
    font.family = ctx$font, cs.family = ctx$cs, eastasia.family = ctx$font,
    hansi.family = ctx$font, font.size = size, bold = bold, italic = italic,
    color = color
  )
}

# Muted colors
.cxf_gray <- "#595959"
.cxf_logic <- "#7F6000"
.cxf_blue <- "#1F3864"

#' Paragraph properties for a role
#'
#' Uses the template style when it exists, else Normal with direct
#' formatting.
#' @noRd
.pp <- function(ctx, role, ...) {
  styles <- c(question = "cxf Question", hint = "cxf Hint",
              logic = "cxf Logic", choice = "cxf Choice", note = "cxf Note",
              banner = "cxf Banner", title = "Title",
              h1 = "heading 1", h2 = "heading 2", h3 = "heading 3",
              normal = "Normal")
  name <- styles[[role]]
  fallback <- switch(
    role,
    question = list(keep_with_next = TRUE, padding.top = 8,
                    padding.bottom = 2),
    hint = list(keep_with_next = TRUE, padding.bottom = 2),
    logic = list(keep_with_next = TRUE, padding.bottom = 2),
    choice = list(padding.left = 21.6, padding.bottom = 2),
    note = list(shading.color = "#F2F2F2", padding.top = 4,
                padding.bottom = 4),
    banner = list(shading.color = "#DEEAF6", keep_with_next = TRUE,
                  padding.top = 6, padding.bottom = 4),
    title = list(text.align = "center", padding.bottom = 8),
    h1 = list(keep_with_next = TRUE, padding.top = 10, padding.bottom = 4),
    h2 = list(keep_with_next = TRUE, padding.top = 8, padding.bottom = 4),
    h3 = list(keep_with_next = TRUE, padding.top = 6, padding.bottom = 4),
    list()
  )
  extra <- list(...)
  if (name %in% ctx$pstyles) {
    args <- c(list(word_style = name), extra)
  } else {
    args <- c(list(word_style = "Normal"),
              utils::modifyList(fallback, extra))
  }
  do.call(officer::fp_par_lite, args)
}

#' Write one block element to the document
#'
#' The body is written through an officer append context (a file of XML)
#' rather than node by node, which keeps rendering time linear in the size
#' of the form.
#' @param doc An `rdocx_append_context` from
#'   `officer::body_append_start_context()`.
#' @param el An `fpar`, a flextable, or a `cxf_wml` table.
#' @return `doc`.
#' @noRd
.emit <- function(doc, el) {
  xml <- if (inherits(el, "cxf_wml")) {
    el$xml
  } else if (inherits(el, "fpar")) {
    .fpar_wml(el)
  } else {
    officer::to_wml(el)
  }
  cat(xml, "\n", sep = "", file = doc$file_con)
  doc
}

#' WML of a paragraph, with the overrides officer leaves out
#'
#' With a paragraph style, officer writes neither `keep_with_next = FALSE`
#' nor a single top or bottom padding, so the style's values win. Both are
#' added here so a block's last paragraph does not keep with the next block.
#' @noRd
.fpar_wml <- function(el) {
  xml <- officer::to_wml(el)
  fp <- el$fp_p
  if (!grepl("<w:pPr>", xml, fixed = TRUE)) {
    return(xml)
  }
  if (isFALSE(fp$keep_with_next) && !grepl("<w:keepNext", xml, fixed = TRUE)) {
    xml <- sub("(<w:pStyle [^>]*/>)", "\\1<w:keepNext w:val=\"0\"/>", xml)
  }
  top <- fp$padding.top %||% NA
  bot <- fp$padding.bottom %||% NA
  if (!grepl("<w:spacing", xml, fixed = TRUE) && (!is.na(top) || !is.na(bot))) {
    sp <- paste0("<w:spacing",
                 if (!is.na(top)) sprintf(" w:before=\"%.0f\"", top * 20),
                 if (!is.na(bot)) sprintf(" w:after=\"%.0f\"", bot * 20),
                 "/>")
    xml <- sub("</w:pPr>", paste0(sp, "</w:pPr>"), xml, fixed = TRUE)
  }
  xml
}

#' Add a paragraph made of runs
#' @noRd
.add_runs <- function(doc, runs, fp_p) {
  .emit(doc, do.call(officer::fpar, c(runs, list(fp_p = fp_p))))
}

#' Add a text paragraph in a role
#' @noRd
.add_text <- function(doc, ctx, text, role, fp, ...) {
  .add_runs(doc, .text_runs(text, fp), .pp(ctx, role, ...))
}

#' A tiny paragraph that stops two tables from merging
#' @noRd
.spacer <- function(doc, ctx) {
  .emit(doc, officer::fpar(
    officer::ftext("", .fp(ctx, size = 3)),
    fp_p = officer::fp_par_lite(word_style = "Normal", padding = 0,
                                keep_with_next = TRUE)
  ))
}

#' Add the elements returned by an answer layout
#' @noRd
.add_elements <- function(doc, ctx, els) {
  prev_table <- FALSE
  for (el in els) {
    is_table <- inherits(el, c("cxf_wml", "flextable"))
    if (is_table && prev_table) doc <- .spacer(doc, ctx)
    doc <- .emit(doc, el)
    prev_table <- is_table
  }
  doc
}

# ---- images ------------------------------------------------------------------

#' Add an image paragraph, or a note when it cannot be embedded
#'
#' Records that are found but have no pixel size (formats other than PNG or
#' JPEG) are not embedded: a distorted picture is worse than a note, so the
#' file name is printed in square brackets after the word "image".
#' @noRd
.add_image <- function(doc, ctx, rec, keep = FALSE) {
  if (is.null(rec)) {
    return(doc)
  }
  usable <- isTRUE(rec$found) && !anyNA(c(rec$width_in, rec$height_in))
  fp_p <- officer::fp_par_lite(word_style = "Normal", padding.bottom = 2,
                               keep_with_next = if (keep) TRUE else NA)
  if (usable) {
    .emit(doc, officer::fpar(
      officer::external_img(rec$path, width = rec$width_in,
                            height = rec$height_in),
      fp_p = fp_p
    ))
  } else {
    .emit(doc, officer::fpar(
      officer::ftext(sprintf("[%s: %s]", ctx$strings$image, rec$file),
                     .fp(ctx, size = ctx$size - 1, italic = TRUE,
                         color = .cxf_gray)),
      fp_p = fp_p
    ))
  }
}

# ---- blocks ------------------------------------------------------------------

#' Logic lines ("Ask if", "Constraint", "Choices shown if") of a block
#' @noRd
.logic_lines <- function(b, ctx) {
  st <- ctx$strings
  out <- character(0)
  add <- function(label, txt) {
    if (!is.null(txt) && length(txt) == 1L && !is.na(txt) && nzchar(txt)) {
      out <<- c(out, paste0(label, ctx$colon, txt))
    }
  }
  add(st$ask_if, b$relevant_text)
  add(st$constraint, b$constraint_text)
  add(st$choice_filter, b$choice_filter_text)
  out
}

#' Add the logic lines; the last one does not keep with the next paragraph
#' @noRd
.add_logic <- function(doc, ctx, lines) {
  n <- length(lines)
  for (k in seq_len(n)) {
    doc <- .add_text(
      doc, ctx, lines[k], "logic",
      .fp(ctx, size = ctx$size - 1, color = .cxf_logic),
      keep_with_next = k < n, padding.top = if (k == 1L) 5 else NA
    )
  }
  doc
}

#' Text of a heading or question: serial, label
#' @noRd
.serial_label <- function(serial, label) {
  lab <- if (is.na(label)) "" else label
  if (!is.na(serial)) paste0(serial, ". ", lab) else lab
}

.render_section <- function(doc, ctx, b) {
  role <- paste0("h", min(3L, max(1L, b$level)))
  hsize <- ctx$size + c(h1 = 2, h2 = 1, h3 = 0)[[role]]
  fp <- .fp(ctx, size = hsize, bold = TRUE, color = .cxf_blue)
  doc <- .add_text(doc, ctx, .serial_label(b$serial, b$label), role, fp)
  lines <- .logic_lines(b, ctx)
  has_img <- !is.null(b$image)
  has_lines <- length(lines) > 0L
  doc <- .add_hint(doc, ctx, b$hint, keep = has_img || has_lines)
  doc <- .add_image(doc, ctx, b$image, keep = has_lines)
  .add_logic(doc, ctx, lines)
}

#' Add a hint paragraph; `keep` says whether something of the same block
#' follows (the last paragraph of a block must not keep with next)
#' @noRd
.add_hint <- function(doc, ctx, hint, keep = TRUE) {
  if (is.null(hint) || is.na(hint) || !nzchar(hint)) {
    return(doc)
  }
  .add_text(doc, ctx, hint, "hint",
            .fp(ctx, size = ctx$size - 1, italic = TRUE, color = .cxf_gray),
            keep_with_next = keep)
}

.render_note <- function(doc, ctx, b) {
  lines <- .logic_lines(b, ctx)
  has_hint <- !is.na(b$hint %||% NA_character_) && nzchar(b$hint)
  has_img <- !is.null(b$image)
  has_lines <- length(lines) > 0L
  doc <- .add_text(doc, ctx, b$label, "note", .fp(ctx, italic = TRUE),
                   keep_with_next = has_hint || has_img || has_lines)
  doc <- .add_hint(doc, ctx, b$hint, keep = has_img || has_lines)
  doc <- .add_image(doc, ctx, b$image, keep = has_lines)
  .add_logic(doc, ctx, lines)
}

.render_question <- function(doc, ctx, b) {
  fp <- .fp(ctx)
  runs <- c(
    list(officer::ftext(if (is.na(b$serial)) "" else paste0(b$serial, ". "),
                        .fp(ctx, bold = TRUE))),
    .text_runs(b$label, fp)
  )
  if (isTRUE(b$required)) {
    runs <- c(runs, list(officer::ftext(
      " *", .fp(ctx, bold = TRUE, color = "#C00000")
    )))
  }
  if (isTRUE(ctx$opts$show_variable_names) && !is.na(b$name)) {
    save_to <- b$save_to %||% NA_character_
    nm <- if (is.na(save_to)) b$name else
      paste0(b$name, " ", ctx$sym$arrow, " ", save_to)
    runs <- c(runs, list(officer::ftext(
      paste0("  [", nm, "]"),
      .fp(ctx, size = ctx$size - 2, color = .cxf_gray)
    )))
  }
  lines <- .logic_lines(b, ctx)
  els <- .answer_elements(b, ctx, has_logic = length(lines) > 0L)
  # the last paragraph or table row of the question never keeps with next
  more <- length(els) > 0L || length(lines) > 0L
  has_hint <- !is.na(b$hint %||% NA_character_) && nzchar(b$hint)
  has_gh <- !is.na(b$guidance_hint %||% NA_character_) &&
    nzchar(b$guidance_hint)
  has_img <- !is.null(b$image)
  doc <- .add_runs(doc, runs, .pp(ctx, "question",
                                  keep_with_next = has_hint || has_gh ||
                                    has_img || more))
  doc <- .add_hint(doc, ctx, b$hint, keep = has_gh || has_img || more)
  doc <- .add_hint(doc, ctx, b$guidance_hint, keep = has_img || more)
  doc <- .add_image(doc, ctx, b$image, keep = more)
  doc <- .add_elements(doc, ctx, els)
  .add_logic(doc, ctx, lines)
}

#' Index of the `section_end` block matching the repeat section at `i`
#' (repeats can nest); `length(blocks) + 1` when it is missing
#' @noRd
.repeat_end <- function(blocks, i) {
  n <- length(blocks)
  depth <- 1L
  j <- i + 1L
  while (j <= n && depth > 0L) {
    if (blocks[[j]]$block == "section" && isTRUE(blocks[[j]]$is_repeat)) {
      depth <- depth + 1L
    } else if (blocks[[j]]$block == "section_end") {
      depth <- depth - 1L
    }
    j <- j + 1L
  }
  if (depth == 0L) j - 1L else n + 1L
}

#' Render a list of blocks, expanding repeats
#' @noRd
.render_blocks <- function(doc, ctx, blocks) {
  n <- length(blocks)
  i <- 1L
  while (i <= n) {
    b <- blocks[[i]]
    switch(
      b$block,
      section = {
        doc <- .render_section(doc, ctx, b)
        g <- .grid_from_section(blocks, i)
        if (!is.null(g)) {
          doc <- .render_grid(doc, ctx, g)
          i <- g$last
        } else if (isTRUE(b$is_repeat)) {
          end <- .repeat_end(blocks, i)
          inner <- if (end - 1L >= i + 1L) blocks[(i + 1L):(end - 1L)] else
            list()
          btxt <- ctx$strings$repeat_for_each
          if (!is.na(b$repeat_text)) btxt <- paste(btxt, b$repeat_text)
          doc <- .add_text(doc, ctx, btxt, "banner",
                           .fp(ctx, size = ctx$size - 1, bold = TRUE,
                               color = .cxf_blue))
          if (.use_roster(inner, ctx$opts)) {
            doc <- .render_roster(doc, ctx, b, inner)
          } else {
            for (cp in seq_len(ctx$opts$repeat_copies)) {
              doc <- .add_text(
                doc, ctx, paste0(ctx$strings$instance_no, " ________"),
                "logic", .fp(ctx, bold = TRUE, color = .cxf_gray),
                padding.top = 6, keep_with_next = TRUE
              )
              doc <- .render_blocks(doc, ctx, inner)
            }
          }
          i <- end
        }
      },
      note = {
        doc <- .render_note(doc, ctx, b)
      },
      question = {
        g <- .grid_from_label(blocks, i)
        if (!is.null(g)) {
          doc <- .render_grid(doc, ctx, g)
          i <- g$last
        } else {
          doc <- .render_question(doc, ctx, b)
        }
      },
      NULL
    )
    i <- i + 1L
  }
  doc
}

# ---- front matter, header, footer -------------------------------------------

#' Legend line runs: the leading symbol is drawn larger
#' @noRd
.legend_runs <- function(line, ctx) {
  sy <- unlist(ctx$sym, use.names = FALSE)
  first <- substr(line, 1L, 1L)
  fp <- .fp(ctx, size = ctx$size - 1)
  if (first %in% sy) {
    c(list(officer::ftext(first, .fp(ctx, size = .symbol_size(first, ctx$size)))),
      .text_runs(substr(line, 2L, nchar(line)), fp))
  } else {
    .text_runs(line, fp)
  }
}

.render_front <- function(doc, ctx) {
  m <- ctx$meta
  st <- ctx$strings
  doc <- .add_text(doc, ctx, m$title, "title",
                   .fp(ctx, size = ctx$size + 5, bold = TRUE))
  parts <- c(
    if (!is.na(m$form_id)) paste0(st$form_id, ctx$colon, m$form_id),
    if (!is.na(m$version)) paste0(st$version, ctx$colon, m$version),
    if (!is.na(m$language)) paste0(st$language, ctx$colon, m$language),
    if (length(m$entities)) {
      paste0(st$entity_list, ctx$colon, paste(m$entities, collapse = ", "))
    }
  )
  if (length(parts) > 0L) {
    doc <- .add_text(doc, ctx, paste(parts, collapse = "   |   "), "hint",
                     .fp(ctx, size = ctx$size - 1, color = .cxf_gray),
                     text.align = "center", keep_with_next = FALSE,
                     padding.bottom = 6)
  }
  leg <- m$legend
  for (k in seq_along(leg)) {
    doc <- .add_runs(
      doc, .legend_runs(leg[k], ctx),
      .pp(ctx, "hint", keep_with_next = FALSE,
          padding.bottom = if (k == length(leg)) 8 else 1)
    )
  }
  doc
}

.header_block <- function(ctx) {
  m <- ctx$meta
  st <- ctx$strings
  fp <- .fp(ctx, size = ctx$size - 2, color = .cxf_gray)
  runs <- list(officer::ftext(sub("\n.*$", "", m$title),
                              .fp(ctx, size = ctx$size - 2,
                                           bold = TRUE, color = .cxf_gray)))
  if (!is.na(m$version)) {
    runs <- c(runs, list(officer::run_tab(),
                         officer::ftext(paste(st$version, m$version), fp)))
  }
  officer::block_list(do.call(officer::fpar, c(runs, list(
    fp_p = officer::fp_par_lite(
      word_style = "Normal",
      tabs = officer::fp_tabs(officer::fp_tab(pos = ctx$tw, style = "right")),
      border.bottom = officer::fp_border(color = "#BFBFBF", width = 0.5),
      padding.bottom = 2
    )
  ))))
}

.footer_block <- function(ctx) {
  m <- ctx$meta
  st <- ctx$strings
  fp <- .fp(ctx, size = ctx$size - 2, color = .cxf_gray)
  left <- paste(
    c(if (!is.na(m$form_id)) m$form_id,
      paste0(st$generated, ctx$colon, format(m$generated, "%Y-%m-%d"))),
    collapse = "   |   "
  )
  runs <- list(
    officer::ftext(left, fp), officer::run_tab(),
    officer::ftext(paste0(st$page, " "), fp),
    officer::run_word_field("PAGE", prop = fp),
    officer::ftext(paste0(" ", st$of, " "), fp),
    officer::run_word_field("NUMPAGES", prop = fp)
  )
  officer::block_list(do.call(officer::fpar, c(runs, list(
    fp_p = officer::fp_par_lite(
      word_style = "Normal",
      tabs = officer::fp_tabs(officer::fp_tab(pos = ctx$tw, style = "right")),
      border.top = officer::fp_border(color = "#BFBFBF", width = 0.5),
      padding.top = 2
    )
  ))))
}

#' Raise complex-script font sizes by 1 pt
#'
#' officer writes the same value for `w:sz` and `w:szCs`; Bengali, Arabic and
#' similar scripts read small at the Latin size, so `w:szCs` is raised.
#' @param x The XML of one document part, as a single string.
#' @noRd
.bump_cs_xml <- function(x) {
  m <- gregexpr("<w:szCs w:val=\"[0-9]+\"/>", x)
  hits <- regmatches(x, m)[[1L]]
  if (length(hits) > 0L) {
    v <- as.integer(gsub("[^0-9]", "", hits)) + 2L
    regmatches(x, m) <- list(sprintf("<w:szCs w:val=\"%d\"/>", v))
  }
  x
}

#' Insert a child element into every `parent` element
#'
#' The child goes before the first existing child named in `before`, which
#' keeps the element order the OOXML schema requires, else at the end.
#' Elements that already have the child are left alone, and self-closing
#' parents are expanded.
#' @param x XML string.
#' @param parent Element name, such as `"w:pPr"`.
#' @param child Child XML, such as `"<w:bidi/>"`.
#' @param before Names of the elements the child must precede.
#' @noRd
.xml_insert_child <- function(x, parent, child, before) {
  x <- gsub(paste0("<", parent, "/>"),
            paste0("<", parent, ">", child, "</", parent, ">"), x,
            fixed = TRUE)
  pat <- paste0("(?s)<", parent, "(\\s[^>]*)?>.*?</", parent, ">")
  m <- gregexpr(pat, x, perl = TRUE)
  els <- regmatches(x, m)[[1L]]
  if (length(els) == 0L) {
    return(x)
  }
  tag <- sub("^<([^ />]+).*$", "\\1", child)
  bpat <- paste0("<(", paste(before, collapse = "|"), ")[\\s/>]")
  els <- vapply(els, function(e) {
    if (grepl(paste0("<", tag, "[\\s/>]"), e, perl = TRUE)) {
      return(e)
    }
    open_end <- regexpr(">", e, fixed = TRUE)
    body <- substring(e, open_end + 1L)
    pos <- regexpr(bpat, body, perl = TRUE)
    if (pos > 0L) {
      body <- paste0(substring(body, 1L, pos - 1L), child,
                     substring(body, pos))
    } else {
      close <- paste0("</", parent, ">")
      body <- paste0(substring(body, 1L, nchar(body) - nchar(close)), child,
                     close)
    }
    paste0(substring(e, 1L, open_end), body)
  }, "", USE.NAMES = FALSE)
  regmatches(x, m) <- list(els)
  x
}

#' Lay out a document part right to left
#'
#' Paragraphs get `w:bidi`, tables `w:bidiVisual`, sections `w:bidi`, and
#' runs that contain Arabic or Hebrew letters `w:rtl`.
#' @param x The XML of one document part, as a single string.
#' @noRd
.rtl_xml <- function(x) {
  # paragraphs without properties get an empty pPr first
  x <- gsub("(<w:p(\\s[^>]*)?>)(?!\\s*<w:pPr[\\s>/])", "\\1<w:pPr></w:pPr>", x,
            perl = TRUE)
  x <- .xml_insert_child(x, "w:pPr", "<w:bidi/>", c(
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange"
  ))
  x <- .xml_insert_child(x, "w:tblPr", "<w:bidiVisual/>", c(
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange"
  ))
  x <- .xml_insert_child(x, "w:sectPr", "<w:bidi/>", c(
    "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange"
  ))
  # runs with right-to-left letters
  m <- gregexpr("(?s)<w:r(\\s[^>]*)?>.*?</w:r>", x, perl = TRUE)
  runs <- regmatches(x, m)[[1L]]
  rtl <- grepl("[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]", runs,
               perl = TRUE)
  if (any(rtl)) {
    runs[rtl] <- vapply(runs[rtl], function(r) {
      if (!grepl("<w:rPr[\\s>/]", r, perl = TRUE)) {
        r <- sub("^(<w:r(\\s[^>]*)?>)", "\\1<w:rPr></w:rPr>", r, perl = TRUE)
      }
      .xml_insert_child(r, "w:rPr", "<w:rtl/>", c(
        "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish",
        "w:oMath"
      ))
    }, "", USE.NAMES = FALSE)
    regmatches(x, m) <- list(runs)
  }
  x
}

#' Patch the XML parts of a saved .docx
#'
#' @param path The .docx file.
#' @param fns Functions, each taking and returning the XML of one part
#'   (document, headers and footers) as a single string.
#' @noRd
.patch_docx <- function(path, fns) {
  if (length(fns) == 0L) {
    return(invisible(path))
  }
  path <- normalizePath(path, winslash = "/", mustWork = TRUE)
  tmp <- tempfile("cxfpatch")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  files <- utils::unzip(path, exdir = tmp)
  xmls <- files[grepl("word/(document|header[0-9]*|footer[0-9]*)\\.xml$",
                      files)]
  for (f in xmls) {
    x <- paste(readLines(f, encoding = "UTF-8", warn = FALSE), collapse = "\n")
    Encoding(x) <- "UTF-8"
    y <- x
    for (fn in fns) y <- fn(y)
    if (!identical(x, y)) {
      con <- file(f, "wb")
      writeBin(charToRaw(enc2utf8(y)), con)
      close(con)
    }
  }
  unlink(path)
  zip::zip(path, list.files(tmp, all.files = TRUE, recursive = TRUE),
           root = tmp, mode = "mirror")
  invisible(path)
}

#' Write the questionnaire as a Word document
#'
#' @param model A `cxf_model` from `.build_model()`.
#' @param opts A `cxf_options` object.
#' @param path Output .docx path.
#' @return `path`, invisibly.
#' @noRd
.render_docx <- function(model, opts, path) {
  template <- opts$template %||%
    system.file("templates", "convertxlsform.docx", package = "convertxlsform")
  doc <- if (nzchar(template)) officer::read_docx(template) else
    officer::read_docx()
  ctx <- .render_ctx(model, opts, doc)

  m <- opts$margins_in
  doc <- officer::body_set_default_section(doc, officer::prop_section(
    page_size = officer::page_size(
      orient = opts$orientation, width = opts$page_width_in,
      height = opts$page_height_in
    ),
    page_margins = officer::page_mar(
      top = m, bottom = m, left = m, right = m,
      header = min(0.5, m / 2), footer = min(0.5, m / 2), gutter = 0
    ),
    type = "continuous",
    header_default = .header_block(ctx),
    footer_default = .footer_block(ctx)
  ))

  body <- officer::body_append_start_context(doc)
  body <- tryCatch({
    body <- .render_front(body, ctx)
    .render_blocks(body, ctx, model$blocks)
  }, error = function(e) {
    close(body$file_con)
    unlink(body$file_path)
    stop(e)
  })
  doc <- officer::body_append_stop_context(body)
  print(doc, target = path)

  .patch_docx(path, c(
    if (!identical(ctx$cs, ctx$font)) list(.bump_cs_xml),
    if (isTRUE(model$meta$rtl)) list(.rtl_xml)
  ))
  invisible(path)
}
