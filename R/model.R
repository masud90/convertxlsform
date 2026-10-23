# The document model: a list of blocks between parsing and rendering.
# ASCII-only source.

#' Combine relevance texts with "and"
#' @noRd
.and_text <- function(parts, strings = NULL) {
  strings <- strings %||% .strings_en()
  parts <- parts[!is.na(parts) & nzchar(parts)]
  if (length(parts) == 0L) {
    return(NA_character_)
  }
  if (length(parts) == 1L) {
    return(parts)
  }
  wrap <- grepl(paste0(" ", strings$or, " "), parts, fixed = TRUE)
  parts[wrap] <- paste0("(", parts[wrap], ")")
  paste(parts, collapse = paste0(" ", strings$and, " "))
}

#' Normalize line breaks to `\n`
#' @noRd
.norm_newlines <- function(x) {
  gsub("\r\n|\r", "\n", x)
}

#' Clean label, hint and note text
#'
#' Line breaks become `\n`, and `${name}` references become `[Q2.2.1]` when
#' the target question is numbered, else `[name]`.
#' @param x Character vector.
#' @param sv The survey data.frame (names and serials).
#' @noRd
.clean_text <- function(x, sv) {
  x <- .norm_newlines(as.character(x))
  has <- !is.na(x) & grepl("${", x, fixed = TRUE)
  for (k in which(has)) {
    m <- gregexpr("\\$\\{[^}]*\\}", x[k])
    refs <- regmatches(x[k], m)[[1L]]
    nm <- trimws(sub("^\\$\\{(.*)\\}$", "\\1", refs))
    j <- match(nm, sv$name)
    ser <- sv$serial[j]
    new <- ifelse(is.na(j) | is.na(ser), paste0("[", nm, "]"),
                  paste0("[Q", ser, "]"))
    regmatches(x[k], m) <- list(new)
  }
  x
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

#' Read the form again in the second language of a bilingual document
#'
#' @return A `cxf_form`, or NULL (with a warning) when `lang2` is not usable.
#' @noRd
.second_language <- function(form, opts) {
  req <- opts$lang2
  if (is.null(req)) {
    return(NULL)
  }
  langs <- form$languages
  if (is.null(langs) || nrow(langs) == 0L) {
    .cxf_warn(c(
      "The form has no language-specific columns; {.arg lang2} {.val {req}} is ignored."
    ))
    return(NULL)
  }
  hit <- .match_language(langs, req)
  if (is.na(hit)) {
    avail <- langs$suffix
    .cxf_warn(c(
      "Second language {.val {req}} not found in the form; it is ignored.",
      i = "Languages found: {.val {avail}}."
    ))
    return(NULL)
  }
  if (identical(tolower(hit), tolower(form$lang$suffix %||% ""))) {
    .cxf_warn("{.arg lang2} is the language already in use; it is ignored.")
    return(NULL)
  }
  form2 <- suppressWarnings(.read_xlsform(form$path, lang = hit,
                                          media_root = form$media_root,
                                          verbose = FALSE))
  if (!isTRUE(opts$nested_serial_numbers)) {
    form2$survey <- .build_structure(form2$survey,
                                     nested_serial_numbers = FALSE)
  }
  form2
}

#' Stack two languages: the second goes on a new line
#'
#' A cell keeps a single text when the second language is blank or the same.
#' @noRd
.stack_text <- function(a, b) {
  a <- as.character(a)
  b <- as.character(b)
  out <- a
  two <- !is.na(a) & !is.na(b) & a != b
  out[two] <- paste0(a[two], "\n", b[two])
  only_b <- is.na(a) & !is.na(b)
  out[only_b] <- b[only_b]
  out
}

#' Merge the text of a second-language form into the first
#' @noRd
.merge_languages <- function(form, form2) {
  for (f in c("label", "hint", "guidance_hint", "constraint_message")) {
    form$survey[[f]] <- .stack_text(form$survey[[f]], form2$survey[[f]])
  }
  if (nrow(form$choices) == nrow(form2$choices)) {
    form$choices$label <- .stack_text(form$choices$label,
                                      form2$choices$label)
  }
  t2 <- form2$settings$form_title
  if (!is.null(t2) && !is.na(t2)) {
    form$settings$form_title <- .stack_text(form$settings$form_title, t2)
  }
  form
}

#' Date, time and datetime entry patterns
#' @noRd
.date_pattern <- function(kind, appearance, strings = NULL) {
  strings <- strings %||% .strings_en()
  app <- tolower(appearance %||% "")
  app <- if (is.na(app)) "" else app
  d <- if (grepl("month-year", app, fixed = TRUE)) {
    paste0("__ / ____  (", strings$fmt_month_year, ")")
  } else if (grepl("\\byear\\b", app)) {
    paste0("____  (", strings$fmt_year, ")")
  } else {
    paste0("__ / __ / ____  (", strings$fmt_date, ")")
  }
  tm <- paste0("__ : __  (", strings$fmt_time, ")")
  switch(kind, date = d, time = tm, datetime = paste0(d, "   ", tm))
}

#' Answer specification for one question
#' @noRd
.build_answer <- function(form, i, opts, constraint = NA_character_,
                          strings = NULL, skip = NULL) {
  strings <- strings %||% .strings_en()
  sv <- form$survey
  k <- sv$kind[i]
  tw <- opts$text_width_in
  switch(
    k,
    select_one = , select_multiple = , rank = {
      sc <- .select_choices(form, i)
      ch <- sc$choices
      n <- nrow(ch)
      ch$label <- .norm_newlines(ch$label)
      imgs <- lapply(ch$image, function(f) .media_record(form, f, .cxf_choice_image_max_in, tw))
      or_other <- isTRUE(sv$or_other[i])
      if (or_other) {
        ch <- data.frame(
          name = c(ch$name, "other"),
          label = c(ch$label, strings$other_specify),
          stringsAsFactors = FALSE
        )
        imgs <- c(imgs, list(NULL))
      }
      ch$image <- imgs
      ch$write_in <- c(rep(FALSE, n), if (or_other) TRUE)
      ch$skip <- if (is.null(skip)) {
        rep(NA_character_, nrow(ch))
      } else {
        vapply(unname(skip[ch$name]), .skip_text, "", strings = strings)
      }
      lay <- .choice_layout(ch$label, sv$appearance[i], opts$max_choices)
      list(
        choices = ch,
        symbol = switch(k, select_one = "circle", select_multiple = "square",
                        "rank"),
        or_other = or_other, n_total = sc$n_total,
        layout = lay$layout, ncol = lay$ncol,
        source = sc$source, source_found = sc$source_found, note = sc$note
      )
    },
    text = list(lines = if (grepl("multiline", tolower(sv$appearance[i] %||% ""),
                                  fixed = TRUE)) 4L else 2L),
    number = list(
      decimal = identical(sv$base_type[i], "decimal"),
      range_text = .simple_range_text(constraint, strings)
    ),
    range = {
      r <- .range_spec(sv$params[[i]])
      r$range_text <- if (is.null(r$values)) {
        sprintf(strings$range_steps, format(r$start), format(r$end),
                format(r$step))
      } else {
        NA_character_
      }
      r
    },
    date = , time = , datetime = list(
      pattern = .date_pattern(k, sv$appearance[i], strings)
    ),
    geo = {
      st <- .type_subtype(sv$base_type[i])
      if (identical(st, "geopoint")) {
        list(subtype = st,
             fields = c(strings$latitude, strings$longitude,
                        strings$accuracy_m))
      } else {
        list(subtype = st, fields = character(0),
             prompt = strings$not_on_paper, reference = strings$reference_no)
      }
    },
    media = list(subtype = .type_subtype(sv$base_type[i]),
                 prompt = strings$not_on_paper,
                 reference = strings$reference_no),
    barcode = list(lines = 1L),
    acknowledge = list(prompt = "OK"),
    hidden = list(calculation = sv$calculation[i]),
    list(lines = 2L)
  )
}

#' Legend lines explaining the symbols used in the form
#' @noRd
.build_legend <- function(blocks, opts, strings = NULL) {
  strings <- strings %||% .strings_en()
  qs <- Filter(function(b) identical(b$block, "question"), blocks)
  sym <- .cxf_symbols()
  has_kind <- function(k) any(vapply(qs, function(b) b$kind == k, NA))
  out <- character(0)
  if (any(vapply(qs, function(b) isTRUE(b$required), NA))) {
    out <- c(out, strings$legend_required)
  }
  if (has_kind("select_one")) {
    out <- c(out, paste(sym$circle, strings$legend_circle))
  }
  if (has_kind("select_multiple")) {
    out <- c(out, paste(sym$square, strings$legend_square))
  }
  if (has_kind("rank")) {
    out <- c(out, paste(sym$rank, strings$legend_rank))
  }
  has_rel <- any(vapply(blocks, function(b) {
    !is.na(b$relevant_text %||% NA_character_)
  }, NA))
  if (isTRUE(opts$include_relevants) && has_rel) {
    out <- c(out, strings$legend_ask_if)
  }
  has_skip <- any(vapply(qs, function(b) {
    any(!is.na(b$answer$choices$skip %||% NA_character_))
  }, NA))
  if (has_skip) {
    out <- c(out, paste(sym$arrow, strings$legend_go_to))
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
  # right-to-left layout follows the first language only
  primary <- .detect_script_font(
    c(form$survey$label, form$survey$hint, form$choices$label),
    opts$font_family
  )
  language <- .language_display(form)
  form2 <- .second_language(form, opts)
  if (!is.null(form2)) {
    language <- paste(language, "/", .language_display(form2))
    form <- .merge_languages(form, form2)
  }
  sv <- form$survey
  n <- nrow(sv)
  .media_reset_missing()

  strings <- .cxf_strings(.form_lang_code(form), opts$strings)
  svn <- opts$show_variable_names
  hx <- function(x, field, fallback = FALSE) {
    .humanize_exprs(x, form, field, rows, strings = strings,
                    show_variable_names = svn, with_fallback = fallback)
  }
  rows <- seq_len(n)
  skips <- if (isTRUE(opts$include_skips)) .skip_targets(form) else
    vector("list", n)
  rel <- if (opts$include_relevants) {
    hx(sv$relevant, "relevant")
  } else {
    rep(NA_character_, n)
  }
  cfilt <- if (opts$include_relevants) {
    hx(sv$choice_filter, "choice_filter")
  } else {
    rep(NA_character_, n)
  }
  # A constraint message replaces the expression. Without a message the
  # humanized expression is shown, unless it could not be humanized and
  # variable names are hidden (then the raw expression is of no use).
  con_msg <- rep(NA_character_, n)
  if (opts$include_constraints) {
    con <- hx(sv$constraint, "constraint", fallback = TRUE)
    has_msg <- !is.na(con$text) & !.is_blank(sv$constraint_message)
    con_msg <- con$text
    con_msg[con$fallback & !svn] <- NA_character_
    con_msg[has_msg] <- .clean_text(sv$constraint_message[has_msg], sv)
  }
  rep_txt <- hx(sv$repeat_count, "repeat_count")
  rep_val <- suppressWarnings(as.numeric(
    ifelse(grepl("^[[:space:]]*[0-9]+[[:space:]]*$", sv$repeat_count),
           sv$repeat_count, NA_character_)
  ))
  label <- .clean_text(sv$label, sv)
  hint <- if (opts$include_hints) .clean_text(sv$hint, sv) else
    rep(NA_character_, n)
  ghint <- if (opts$include_hints) .clean_text(sv$guidance_hint, sv) else
    rep(NA_character_, n)
  has_label <- !.is_blank(sv$label)
  tw <- opts$text_width_in
  img <- function(i, h = .cxf_question_image_max_in) .media_record(form, sv$image[i], h, tw)

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
    own_rel <- .and_text(c(inherited[!is.na(inherited)], rel[i]), strings)

    if (k %in% c("begin_group", "begin_repeat")) {
      shown <- k == "begin_repeat" || has_label[i]
      if (shown) {
        add(c(list(block = "section"), common, list(
          level = min(3L, n_vis + 1L), label = label[i], hint = hint[i],
          is_repeat = k == "begin_repeat", appearance = sv$appearance[i],
          repeat_text = if (k == "begin_repeat") rep_txt[i] else NA_character_,
          repeat_count_value = if (k == "begin_repeat") rep_val[i] else
            NA_real_,
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
          level = min(3L, n_vis), label = label[i], hint = hint[i],
          image = img(i), relevant_text = own_rel
        )))
      }
    } else if (k == "hidden" && !opts$include_hidden) {
      next
    } else {
      is_hidden <- k == "hidden"
      add(c(list(block = "question"), common, list(
        level = min(3L, n_vis), label = label[i], hint = hint[i],
        guidance_hint = ghint[i], appearance = sv$appearance[i],
        required = isTRUE(sv$required[i]), kind = k,
        subtype = if (is_hidden) sv$base_type[i] else
          .type_subtype(sv$base_type[i]),
        relevant_text = own_rel, constraint_text = con_msg[i],
        choice_filter_text = cfilt[i], image = img(i),
        save_to = sv$save_to[i] %||% NA_character_,
        answer = .build_answer(form, i, opts, sv$constraint[i], strings,
                               skip = skips[[i]])
      ), if (is_hidden) list(calculation = sv$calculation[i])))
    }
  }

  title <- opts$title
  if (is.null(title)) title <- form$settings$form_title
  if (is.null(title) || .is_blank(title)) {
    title <- tools::file_path_sans_ext(basename(form$path))
  }
  # Fonts for complex scripts come from the labels in the language in use.
  cs <- .detect_script_font(
    c(sv$label, sv$hint, form$choices$label), opts$font_family
  )
  if (!is.null(opts$font_family_cs)) {
    cs$font_family_cs <- opts$font_family_cs
  }
  meta <- list(
    title = title,
    form_id = form$settings$form_id %||% NA_character_,
    version = form$settings$version %||% NA_character_,
    language = language,
    generated = Sys.Date(),
    legend = .build_legend(blocks, opts, strings),
    entities = unique(form$entities$list_name %||% character(0)),
    font_family_cs = cs$font_family_cs,
    script = cs$script,
    rtl = if (is.null(opts$rtl)) primary$rtl else opts$rtl,
    strings = strings
  )
  .media_warn_missing()
  structure(list(meta = meta, blocks = blocks), class = "cxf_model")
}
