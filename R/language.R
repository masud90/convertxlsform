# Language detection and per-field language resolution.
# ASCII-only source.

.lang_fields <- function() {
  c("label", "hint", "guidance_hint", "required_message",
    "constraint_message", "image", "audio", "video")
}

#' Normalize sheet headers
#'
#' Trims and lower-cases the field part of a header, keeps the language
#' suffix as written, and maps `media::image` (and audio, video) to the plain
#' field, so `media::image::English (en)` becomes `image::English (en)`.
#'
#' @param x Character vector of raw headers.
#' @return Character vector of normalized headers.
#' @noRd
.norm_header <- function(x) {
  x <- as.character(x)
  out <- x
  ok <- !is.na(x)
  y <- trimws(x[ok])
  y <- sub("^media[[:space:]]*::[[:space:]]*(image|audio|video)", "\\1", y,
           ignore.case = TRUE, perl = TRUE)
  has <- grepl("::", y, fixed = TRUE)
  field <- tolower(trimws(sub("::.*$", "", y)))
  suffix <- trimws(sub("^.*?::", "", y, perl = TRUE))
  field <- gsub("[[:space:]]+", " ", field)
  field[field == "list name"] <- "list_name"
  res <- ifelse(has, paste0(field, "::", suffix), field)
  res[has & !nzchar(suffix)] <- field[has & !nzchar(suffix)]
  out[ok] <- res
  out
}

#' Split normalized headers into field and suffix
#' @noRd
.split_header <- function(x) {
  has <- grepl("::", x, fixed = TRUE)
  data.frame(
    field = sub("::.*$", "", x),
    suffix = ifelse(has, sub("^.*?::", "", x, perl = TRUE), NA_character_),
    stringsAsFactors = FALSE
  )
}

#' Parse language suffixes into suffix, code and name
#' @noRd
.parse_lang_suffix <- function(suffix) {
  suffix <- as.character(suffix)
  n <- length(suffix)
  code <- rep(NA_character_, n)
  name <- rep(NA_character_, n)
  paren <- grepl("^.*\\([^()]+\\)[[:space:]]*$", suffix)
  if (any(paren)) {
    name[paren] <- trimws(sub("[[:space:]]*\\([^()]+\\)[[:space:]]*$", "",
                              suffix[paren]))
    code[paren] <- trimws(sub("^.*\\(([^()]+)\\)[[:space:]]*$", "\\1",
                              suffix[paren]))
    name[paren & !nzchar(name)] <- NA_character_
  }
  is_code <- !paren &
    grepl("^[A-Za-z]{2,3}([-_][A-Za-z0-9]{2,4})?$", suffix)
  code[is_code] <- suffix[is_code]
  is_name <- !paren & !is_code
  name[is_name] <- suffix[is_name]
  data.frame(suffix = suffix, code = code, name = name,
             stringsAsFactors = FALSE)
}

#' Detect languages from normalized survey and choices headers
#' @noRd
.detect_languages <- function(...) {
  hdr <- unlist(list(...), use.names = FALSE)
  hdr <- hdr[!is.na(hdr)]
  sp <- .split_header(hdr)
  keep <- sp$field %in% .lang_fields() & !is.na(sp$suffix) & nzchar(sp$suffix)
  suf <- sp$suffix[keep]
  suf <- suf[!duplicated(tolower(suf))]
  out <- .parse_lang_suffix(suf)
  need <- is.na(out$code) & !is.na(out$name)
  out$code[need] <- unname(.lang_name_codes()[.fold_name(out$name[need])])
  out
}

#' Lower-case a language name and drop accents that matter here
#' @noRd
.fold_name <- function(x) {
  x <- tolower(trimws(x))
  x <- gsub("\u00e7", "c", x, fixed = TRUE)
  x <- gsub("\u00e9", "e", x, fixed = TRUE)
  x <- gsub("\u00f1", "n", x, fixed = TRUE)
  x <- gsub("\u00ea", "e", x, fixed = TRUE)
  x
}

#' ISO 639-1 codes for common language names
#'
#' Used only to fill a missing code for headers such as `label::French`.
#' Languages without an ISO 639-1 code (for example Dinka, Nuer, Dari) are left
#' out; Dari is listed under its ISO 639-3 code `prs`.
#' @noRd
.lang_name_codes <- function() {
  c(english = "en", french = "fr", francais = "fr", spanish = "es",
    espanol = "es", portuguese = "pt", arabic = "ar", swahili = "sw",
    kiswahili = "sw", somali = "so", amharic = "am", tigrinya = "ti",
    oromo = "om", oromiffa = "om", bengali = "bn", bangla = "bn",
    burmese = "my", hindi = "hi", urdu = "ur", pashto = "ps", dari = "prs",
    persian = "fa", farsi = "fa", ukrainian = "uk", russian = "ru",
    turkish = "tr", kinyarwanda = "rw", kirundi = "rn", lingala = "ln",
    hausa = "ha")
}

#' Match a requested language against the detected ones
#'
#' @return The matching suffix, or NA.
#' @noRd
.match_language <- function(langs, requested) {
  if (is.null(requested) || length(requested) != 1L || is.na(requested) ||
      nrow(langs) == 0L) {
    return(NA_character_)
  }
  r <- tolower(trimws(requested))
  hit <- tolower(langs$suffix) == r |
    (!is.na(langs$code) & tolower(langs$code) == r) |
    (!is.na(langs$name) & tolower(langs$name) == r)
  if (any(hit)) langs$suffix[which(hit)[1L]] else NA_character_
}

#' Decide which language to use
#'
#' @return list(requested, suffix, mode). Warns when a requested language is
#'   not found.
#' @noRd
.resolve_language <- function(langs, requested, default_language) {
  req <- if (is.null(requested)) NA_character_ else requested
  if (nrow(langs) == 0L) {
    if (!is.na(req)) {
      .cxf_warn(c(
        "The form has no language-specific columns; {.arg lang} {.val {req}} is ignored.",
        i = "Using the plain {.field label} columns."
      ))
    }
    return(list(requested = req, suffix = NA_character_, mode = "single"))
  }
  def <- .match_language(langs, default_language)
  fallback <- if (!is.na(def)) def else langs$suffix[1L]
  if (is.na(req)) {
    return(list(requested = req, suffix = fallback, mode = "multilingual"))
  }
  hit <- .match_language(langs, req)
  if (!is.na(hit)) {
    return(list(requested = req, suffix = hit, mode = "multilingual"))
  }
  avail <- langs$suffix
  .cxf_warn(c(
    "Language {.val {req}} not found in the form.",
    i = "Languages found: {.val {avail}}.",
    i = "Using {.val {fallback}} instead."
  ))
  list(requested = req, suffix = fallback, mode = "fallback")
}

#' Resolve one field for the chosen language
#'
#' Per cell: `field::<suffix>`, then plain `field`, then the default
#' language's column.
#'
#' @param df Data.frame with normalized headers.
#' @return Character vector with one value per row.
#' @noRd
.resolve_field <- function(df, field, suffix, default_suffix = NA_character_) {
  n <- nrow(df)
  out <- rep(NA_character_, n)
  if (n == 0L || ncol(df) == 0L) {
    return(out)
  }
  sp <- .split_header(names(df))
  pick <- function(sfx) {
    if (is.na(sfx)) {
      i <- which(sp$field == field & is.na(sp$suffix))
    } else {
      i <- which(sp$field == field & !is.na(sp$suffix) &
                   tolower(sp$suffix) == tolower(sfx))
    }
    if (length(i) == 0L) rep(NA_character_, n) else as.character(df[[i[1L]]])
  }
  cands <- list(pick(suffix), pick(NA_character_), pick(default_suffix))
  if (is.na(suffix)) cands <- cands[-1L]
  for (v in cands) {
    todo <- .is_blank(out)
    if (!any(todo)) break
    out[todo] <- v[todo]
  }
  out[.is_blank(out)] <- NA_character_
  out
}

#' List the languages of an XLSForm
#'
#' Finds the languages used in the `label::<language>` style columns of the
#' survey and choices sheets, for example `label::English (en)`, `label::en`
#' or `label::English`.
#'
#' @param xlsform Path to an XLSForm (`.xlsx` or `.xls`).
#' @return A data.frame with one row per language and columns `suffix` (the
#'   text after `::`), `code`, `name` and `default` (TRUE for the language used
#'   when no `lang` is requested: the `default_language` of the settings sheet
#'   if it matches, otherwise the first language). The data.frame has no rows
#'   for a single-language form.
#' @examples
#' path <- system.file("extdata", "multilang", "multilang.xlsx",
#'                     package = "convertxlsform")
#' xlsform_languages(path)
#' @export
xlsform_languages <- function(xlsform) {
  if (inherits(xlsform, "cxf_form")) {
    langs <- xlsform$languages
    def <- xlsform$settings$default_language
  } else {
    raw <- .read_raw_sheets(.check_xlsform_path(xlsform))
    langs <- .detect_languages(names(raw$survey), names(raw$choices))
    def <- .settings_from_raw(raw$settings, NA_character_,
                              NA_character_)$default_language
  }
  m <- .match_language(langs, def)
  idx <- if (!is.na(m)) match(m, langs$suffix) else 1L
  langs$default <- if (nrow(langs) > 0L) seq_len(nrow(langs)) == idx else
    logical(0)
  rownames(langs) <- NULL
  langs
}
