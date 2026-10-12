# The document model: a list of blocks between parsing and rendering.
# ASCII-only source.

#' Combine relevance texts with "and"
#' @noRd
.and_text <- function(parts) {
  parts <- parts[!is.na(parts) & nzchar(parts)]
  if (length(parts) == 0L) {
    return(NA_character_)
  }
  if (length(parts) == 1L) {
    return(parts)
  }
  wrap <- grepl("\\bor\\b", parts)
  parts[wrap] <- paste0("(", parts[wrap], ")")
  paste(parts, collapse = " and ")
}

#' Display name of the language in use (NA for single-language forms)
#' @noRd
.language_display <- function(form) {
  sfx <- form$lang$suffix
  if (is.null(sfx) || is.na(sfx)) {
    return(NA_character_)
  }
  langs <- form$languages
  i <- match(tolower(sfx), tolower(langs$suffix))
  if (is.na(i)) {
    return(sfx)
  }
  nm <- langs$name[i]
  if (is.na(nm)) nm <- langs$code[i]
  if (is.na(nm)) sfx else nm
}

#' Date, time and datetime entry patterns
#' @noRd
.date_pattern <- function(kind, appearance) {
  app <- tolower(appearance %||% "")
  app <- if (is.na(app)) "" else app
  d <- if (grepl("month-year", app, fixed = TRUE)) {
    "__ / ____  (mm/yyyy)"
  } else if (grepl("\\byear\\b", app)) {
    "____  (yyyy)"
  } else {
    "__ / __ / ____  (dd/mm/yyyy)"
  }
  tm <- "__ : __  (hh:mm)"
  switch(kind, date = d, time = tm, datetime = paste0(d, "   ", tm))
}

#' Answer specification for one question
#' @noRd
.build_answer <- function(form, i, opts, constraint = NA_character_) {
  sv <- form$survey
  k <- sv$kind[i]
  tw <- opts$text_width_in
  switch(
    k,
    select_one = , select_multiple = , rank = {
      sc <- .select_choices(form, i, opts$max_choices)
      ch <- sc$choices
      ch$image <- lapply(ch$image, function(f) {
        .media_record(form, f, 1, tw)
      })
      list(
        choices = ch,
        symbol = switch(k, select_one = "circle", select_multiple = "square",
                        "rank"),
        or_other = isTRUE(sv$or_other[i]), n_total = sc$n_total,
        truncated = sc$truncated, source = sc$source,
        source_found = sc$source_found, note = sc$note
      )
    },
    text = list(lines = if (grepl("multiline", tolower(sv$appearance[i] %||% ""),
                                  fixed = TRUE)) 4L else 2L),
    number = list(
      decimal = identical(sv$base_type[i], "decimal"),
      range_text = .simple_range_text(constraint)
    ),
    range = .range_spec(sv$params[[i]]),
    date = , time = , datetime = list(
      pattern = .date_pattern(k, sv$appearance[i])
    ),
    geo = {
      st <- .type_subtype(sv$base_type[i])
      list(subtype = st,
           fields = if (identical(st, "geopoint")) {
             c("Latitude", "Longitude", "Altitude", "Accuracy")
           } else {
             "List of points"
           })
    },
    media = {
      st <- .type_subtype(sv$base_type[i])
      list(subtype = st, prompt = switch(
        st,
        image = "[Photo to be taken]",
        audio = "[Audio to be recorded]",
        video = "[Video to be recorded]",
        file = "[File to be attached]",
        "background-audio" = "[Background audio is recorded automatically]",
        "[File to be attached]"
      ))
    },
    barcode = list(lines = 1L),
    acknowledge = list(prompt = "OK"),
    hidden = list(calculation = sv$calculation[i]),
    list(lines = 2L)
  )
}

#' Legend lines explaining the symbols used in the form
#' @noRd
.build_legend <- function(blocks, opts) {
  qs <- Filter(function(b) identical(b$block, "question"), blocks)
  sym <- .cxf_symbols()
  has_kind <- function(k) any(vapply(qs, function(b) b$kind == k, NA))
  out <- character(0)
  if (any(vapply(qs, function(b) isTRUE(b$required), NA))) {
    out <- c(out, "* Required question")
  }
  if (has_kind("select_one")) {
    out <- c(out, paste(sym$circle, "Choose one answer"))
  }
  if (has_kind("select_multiple")) {
    out <- c(out, paste(sym$square, "Choose all answers that apply"))
  }
  has_rel <- any(vapply(blocks, function(b) {
    !is.na(b$relevant_text %||% NA_character_)
  }, NA))
  if (isTRUE(opts$include_relevants) && has_rel) {
    out <- c(out, "Ask if: ask the question only when the condition is met")
  }
  out
}

#' Build the document model
#'
#' @param form A `cxf_form`, numbered with `nested_serial_numbers = TRUE`.
#' @param opts A `cxf_options` object (see `.cxf_options()`).
#' @return A `cxf_model`: `meta` and `blocks` (model spec section 6). Every
#'   section, question and note block also carries `level`: for a section the
#'   heading level (1 to 3, from visible enclosing sections only); for a
#'   question or note the number of visible enclosing sections (0 when not
#'   inside one, capped at 3).
#' @noRd
.build_model <- function(form, opts) {
  if (!isTRUE(opts$nested_serial_numbers)) {
    form$survey <- .build_structure(form$survey, nested_serial_numbers = FALSE)
  }
  sv <- form$survey
  n <- nrow(sv)
  .media_reset_missing()

  rows <- seq_len(n)
  rel <- if (opts$include_relevants) {
    .humanize_exprs(sv$relevant, form, "relevant", rows)
  } else {
    rep(NA_character_, n)
  }
  cfilt <- if (opts$include_relevants) {
    .humanize_exprs(sv$choice_filter, form, "choice_filter", rows)
  } else {
    rep(NA_character_, n)
  }
  con <- if (opts$include_constraints) {
    .humanize_exprs(sv$constraint, form, "constraint", rows)
  } else {
    rep(NA_character_, n)
  }
  con_msg <- ifelse(!is.na(con) & !is.na(sv$constraint_message),
                    paste0(con, " (", sv$constraint_message, ")"), con)
  rep_txt <- .humanize_exprs(sv$repeat_count, form, "repeat_count", rows)
  hint <- if (opts$include_hints) sv$hint else rep(NA_character_, n)
  has_label <- !.is_blank(sv$label)
  tw <- opts$text_width_in
  img <- function(i, h = 2) .media_record(form, sv$image[i], h, tw)

  blocks <- list()
  add <- function(b) blocks[[length(blocks) + 1L]] <<- b

  vis <- logical(0)         # per open begin_* row: shown as a section?
  inherited <- character(0) # per open begin_* row: relevance of a hidden group
  open_idx <- integer(0)    # per open begin_* row: survey index

  for (i in rows) {
    k <- sv$kind[i]
    n_vis <- sum(vis)
    common <- list(row = sv$row[i], depth = sv$depth[i], serial = sv$serial[i],
                   name = sv$name[i])
    own_rel <- .and_text(c(inherited[!is.na(inherited)], rel[i]))

    if (k %in% c("begin_group", "begin_repeat")) {
      shown <- k == "begin_repeat" || has_label[i]
      if (shown) {
        add(c(list(block = "section"), common, list(
          level = min(3L, n_vis + 1L), label = sv$label[i], hint = hint[i],
          is_repeat = k == "begin_repeat",
          repeat_text = if (k == "begin_repeat") rep_txt[i] else NA_character_,
          relevant_text = own_rel, image = img(i)
        )))
      }
      vis <- c(vis, shown)
      open_idx <- c(open_idx, i)
      inherited <- c(inherited, if (!shown) rel[i] else NA_character_)
    } else if (k %in% c("end_group", "end_repeat")) {
      if (k == "end_repeat" && length(vis) > 0L) {
        # the closing row carries the serial and name of its repeat
        b <- open_idx[length(open_idx)]
        add(c(list(block = "section_end"),
              list(row = sv$row[i], depth = sv$depth[i], serial = sv$serial[b],
                   name = sv$name[b]),
              list(level = min(3L, n_vis))))
      }
      if (length(vis) > 0L) {
        vis <- vis[-length(vis)]
        inherited <- inherited[-length(inherited)]
        open_idx <- open_idx[-length(open_idx)]
      }
    } else if (k == "note") {
      if (has_label[i] || !.is_blank(sv$image[i])) {
        add(c(list(block = "note"), common, list(
          level = min(3L, n_vis), label = sv$label[i], hint = hint[i],
          image = img(i), relevant_text = own_rel
        )))
      }
    } else if (k == "hidden" && !opts$include_hidden) {
      next
    } else {
      is_hidden <- k == "hidden"
      add(c(list(block = "question"), common, list(
        level = min(3L, n_vis), label = sv$label[i], hint = hint[i],
        required = isTRUE(sv$required[i]), kind = k,
        subtype = if (is_hidden) sv$base_type[i] else
          .type_subtype(sv$base_type[i]),
        relevant_text = own_rel, constraint_text = con_msg[i],
        choice_filter_text = cfilt[i], image = img(i),
        answer = .build_answer(form, i, opts, sv$constraint[i])
      ), if (is_hidden) list(calculation = sv$calculation[i])))
    }
  }

  title <- opts$title
  if (is.null(title)) title <- form$settings$form_title
  if (is.null(title) || .is_blank(title)) {
    title <- tools::file_path_sans_ext(basename(form$path))
  }
  meta <- list(
    title = title,
    form_id = form$settings$form_id %||% NA_character_,
    version = form$settings$version %||% NA_character_,
    language = .language_display(form),
    generated = Sys.Date(),
    legend = .build_legend(blocks, opts)
  )
  .media_warn_missing()
  structure(list(meta = meta, blocks = blocks), class = "cxf_model")
}
