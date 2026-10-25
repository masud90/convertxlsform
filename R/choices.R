# Choice lists: the choices sheet, external CSV and XML files, range values and
# truncation. ASCII-only source.

#' Suffix of the default language (NA for single-language forms)
#' @noRd
.default_suffix <- function(form) {
  langs <- form$languages
  if (is.null(langs) || nrow(langs) == 0L) {
    return(NA_character_)
  }
  d <- .match_language(langs, form$settings$default_language)
  if (is.na(d)) langs$suffix[1L] else d
}

#' Choices of one list from the choices sheet
#'
#' @param form A `cxf_form`.
#' @param list_name The list name.
#' @return A data.frame with `name`, `label`, `image` (file name as written,
#'   NA if none) and `row`, in sheet order. A blank label falls back to the
#'   choice name. Zero rows if the list does not exist.
#' @noRd
.choice_list <- function(form, list_name) {
  ch <- form$choices
  keep <- !is.na(ch$list_name) & !is.na(list_name) & ch$list_name == list_name
  out <- ch[keep, c("name", "label", "image", "row"), drop = FALSE]
  blank <- .is_blank(out$label)
  out$label[blank] <- out$name[blank]
  rownames(out) <- NULL
  out
}

#' Empty choices data.frame
#' @noRd
.empty_choices <- function() {
  data.frame(name = character(0), label = character(0),
             image = character(0), stringsAsFactors = FALSE)
}

#' Find an external choices file under media_root
#'
#' Tries the name as given (relative to `media_root`, or absolute), then its
#' basename under `media_root`.
#' @return The path, or NA.
#' @noRd
.find_media_file <- function(form, file) {
  if (is.null(file) || length(file) != 1L || .is_blank(file)) {
    return(NA_character_)
  }
  root <- form$media_root
  as_given <- if (grepl("^([A-Za-z]:)?[/\\\\]", file) ||
                  startsWith(file, "~")) {
    path.expand(file)
  } else {
    file.path(root, file)
  }
  cand <- unique(c(as_given, file.path(root, basename(file))))
  ok <- file.exists(cand) & !dir.exists(cand)
  if (!any(ok)) {
    return(NA_character_)
  }
  normalizePath(cand[ok][1L], winslash = "/", mustWork = FALSE)
}

#' Choices for a `*_from_file` question
#'
#' A `.csv` or `.xml` file under `media_root` is read; the label column is
#' chosen with the same rules as the survey sheet (`label::<language>`, then
#' `label`, then the default language, then any label column, then the `name`
#' column). Other formats (such as `.geojson`) are located but not parsed.
#'
#' @param form A `cxf_form`.
#' @param file File name as written in the type.
#' @return A list: `choices` (name, label, image), `found` (usable file),
#'   `note` (NA or a short reason), `path` (NA if not located).
#' @noRd
.file_choices <- function(form, file) {
  res <- function(choices = .empty_choices(), found = FALSE,
                  note = NA_character_, path = NA_character_) {
    list(choices = choices, found = found, note = note, path = path)
  }
  path <- .find_media_file(form, file)
  if (is.na(path)) {
    .media_env$missing <- c(.media_env$missing, file)
    return(res(note = "File not found."))
  }
  ext <- tolower(tools::file_ext(path))
  if (!(ext %in% c("csv", "xml"))) {
    return(res(found = TRUE, path = path,
               note = "Only CSV and XML choice files are read."))
  }
  df <- tryCatch(
    if (ext == "csv") {
      suppressWarnings(utils::read.csv(
        path, fileEncoding = "UTF-8-BOM", colClasses = "character",
        check.names = FALSE, stringsAsFactors = FALSE,
        na.strings = character(0)
      ))
    } else {
      .read_xml_items(path)
    },
    error = function(e) NULL
  )
  if (is.null(df) || ncol(df) == 0L) {
    return(res(path = path, note = "File could not be read."))
  }
  names(df) <- .norm_header(names(df))
  df[] <- lapply(df, function(v) {
    v <- trimws(v)
    v[!nzchar(v)] <- NA_character_
    v
  })
  nm <- if ("name" %in% names(df)) df[["name"]] else df[[1L]]
  suffix <- form$lang$suffix %||% NA_character_
  label <- .resolve_field(df, "label", suffix, .default_suffix(form))
  if (any(.is_blank(label))) {
    sp <- .split_header(names(df))
    for (j in which(sp$field == "label")) {
      todo <- .is_blank(label)
      label[todo] <- df[[j]][todo]
    }
  }
  label[.is_blank(label)] <- nm[.is_blank(label)]
  image <- .resolve_field(df, "image", suffix, .default_suffix(form))
  choices <- data.frame(name = nm, label = label, image = image,
                        stringsAsFactors = FALSE)
  choices <- choices[!is.na(choices$name), , drop = FALSE]
  rownames(choices) <- NULL
  res(choices = choices, found = TRUE, path = path)
}

#' Replace XML character and entity references
#' @noRd
.xml_unescape <- function(x) {
  num <- function(x, pat, base) {
    m <- gregexpr(pat, x, perl = TRUE)
    regmatches(x, m) <- lapply(regmatches(x, m), function(r) {
      cp <- strtoi(sub(pat, "\\1", r, perl = TRUE), base)
      vapply(cp, function(k) {
        if (is.na(k)) "" else intToUtf8(k)
      }, "")
    })
    x
  }
  x <- num(x, "&#([0-9]+);", 10L)
  x <- num(x, "&#[xX]([0-9A-Fa-f]+);", 16L)
  x <- gsub("&lt;", "<", x, fixed = TRUE)
  x <- gsub("&gt;", ">", x, fixed = TRUE)
  x <- gsub("&quot;", "\"", x, fixed = TRUE)
  x <- gsub("&apos;", "'", x, fixed = TRUE)
  gsub("&amp;", "&", x, fixed = TRUE)
}

#' Text of an XML fragment: CDATA kept, tags dropped, references replaced
#' @noRd
.xml_text <- function(x) {
  if (!nzchar(x)) {
    return("")
  }
  parts <- strsplit(x, "<![CDATA[", fixed = TRUE)[[1L]]
  out <- .xml_unescape(gsub("<[^>]*>", "", parts[1L]))
  for (p in parts[-1L]) {
    k <- regexpr("]]>", p, fixed = TRUE)
    if (k < 0L) stop("unclosed CDATA", call. = FALSE)
    rest <- substring(p, k + 3L)
    out <- paste0(out, substr(p, 1L, k - 1L),
                  .xml_unescape(gsub("<[^>]*>", "", rest)))
  }
  out
}

#' Read the items of an ODK XML choice file
#'
#' The file has the shape `<root><item><name>a</name><label>A</label></item>
#' ...</root>`. Each `item` element becomes a row and each child element a
#' column (by element name), so the same label rules as for CSV files apply.
#' Base R only: comments, the XML declaration and CDATA sections are handled;
#' namespaces and nested elements inside a child are not.
#'
#' @param path Path to the file.
#' @return A data.frame of character columns, or NULL when there are no items.
#' @noRd
.read_xml_items <- function(path) {
  x <- paste(readLines(path, encoding = "UTF-8", warn = FALSE),
             collapse = "\n")
  x <- sub("^\ufeff", "", x)
  # (?s): the lazy matches may span lines
  x <- gsub("(?s)<!--.*?-->", "", x, perl = TRUE)
  x <- gsub("(?s)<\\?.*?\\?>", "", x, perl = TRUE)
  items <- regmatches(x, gregexpr(
    "(?s)<item\\b[^>]*?(?:/>|>.*?</item[[:space:]]*>)", x, perl = TRUE
  ))[[1L]]
  if (length(items) == 0L) {
    return(NULL)
  }
  child_pat <- paste0("(?s)<([A-Za-z_][A-Za-z0-9_.:-]*)\\b[^>]*?",
                      "(?:/>|>(.*?)</\\1[[:space:]]*>)")
  rows <- lapply(items, function(it) {
    body <- sub("^<item\\b[^>]*>", "", it, perl = TRUE)
    body <- sub("</item[[:space:]]*>$", "", body, perl = TRUE)
    kids <- regmatches(body, gregexpr(child_pat, body, perl = TRUE))[[1L]]
    if (length(kids) == 0L) {
      return(character(0))
    }
    whole <- sub("^\\(\\?s\\)", "(?s)^", paste0(child_pat, "$"))
    nms <- sub(whole, "\\1", kids, perl = TRUE)
    vals <- sub(whole, "\\2", kids, perl = TRUE)
    vals <- vapply(vals, .xml_text, "", USE.NAMES = FALSE)
    stats::setNames(vals, nms)
  })
  cols <- unique(unlist(lapply(rows, names), use.names = FALSE))
  if (length(cols) == 0L) {
    return(NULL)
  }
  out <- lapply(cols, function(cl) {
    vapply(rows, function(r) {
      if (cl %in% names(r)) r[[cl]] else NA_character_
    }, "")
  })
  names(out) <- cols
  as.data.frame(out, stringsAsFactors = FALSE, optional = TRUE)
}

#' Layout of a choice list on paper
#'
#' `"list"` for up to 7 choices; `"columns"` for 8 up to `max_choices`, in 2
#' or 3 columns (3 when every label has at most 25 characters); `"code_box"`
#' for more than `max_choices`, or when the appearance contains `minimal` or
#' `autocomplete` and there are more than 7 choices: a write-in code box plus
#' a compact code table of all choices.
#'
#' @param labels Character vector of choice labels (including any "other" row).
#' @param appearance The question's appearance (may be NA).
#' @param max_choices Largest list still printed in full.
#' @return list(layout, ncol)
#' @noRd
.choice_layout <- function(labels, appearance = NA_character_,
                           max_choices = 30) {
  n <- length(labels)
  app <- tolower(if (is.null(appearance) || is.na(appearance)) "" else appearance)
  compact <- grepl("minimal", app, fixed = TRUE) ||
    grepl("autocomplete", app, fixed = TRUE)
  short <- all(nchar(labels[!is.na(labels)]) <= 25L)
  ncol3 <- if (short) 3L else 2L
  if (n > max_choices || (compact && n > 7L)) {
    list(layout = "code_box", ncol = ncol3)
  } else if (n <= 7L) {
    list(layout = "list", ncol = 1L)
  } else {
    list(layout = "columns", ncol = ncol3)
  }
}

#' Choices and source details for a select or rank question
#'
#' @param form A `cxf_form`.
#' @param i Survey row index of the question.
#' @return list(choices, n_total, source, source_found, note). `n_total` is
#'   the number of choices of the list (before any "other" row is added).
#' @noRd
.select_choices <- function(form, i) {
  sv <- form$survey
  bt <- sv$base_type[i]
  arg <- sv$type_arg[i]
  if (grepl("_from_file$", bt)) {
    fc <- .file_choices(form, arg)
    ch <- fc$choices
    source <- arg
    found <- fc$found
    note <- fc$note
  } else if (grepl("_external$", bt)) {
    ch <- .empty_choices()
    source <- arg
    found <- FALSE
    note <- "External choices are not read."
  } else {
    ch <- .choice_list(form, arg)[, c("name", "label", "image"), drop = FALSE]
    source <- "choices"
    found <- TRUE
    note <- NA_character_
  }
  list(choices = ch, n_total = nrow(ch), source = source, source_found = found,
       note = note)
}

#' Start, end, step and the values of a range question
#'
#' Defaults are start 1, end 10, step 1. `values` is NULL when there are more
#' than 20 values or the range is empty.
#'
#' @param params Named list from the `parameters` column.
#' @return list(start, end, step, values)
#' @noRd
.range_spec <- function(params) {
  num <- function(key, default) {
    v <- params[[key]]
    if (is.numeric(v) && length(v) == 1L && is.finite(v)) v else default
  }
  start <- num("start", 1)
  end <- num("end", 10)
  step <- num("step", 1)
  if (step <= 0) {
    step <- 1
  }
  count <- if (end >= start) floor((end - start) / step + 1e-9) + 1 else 0
  values <- if (count >= 1 && count <= 20) {
    start + step * (seq_len(count) - 1)
  } else {
    NULL
  }
  list(start = start, end = end, step = step, values = values)
}
