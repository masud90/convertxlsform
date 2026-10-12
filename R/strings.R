# Connective strings used by the humanizer, the model and the renderer, with
# built-in English and French tables. ASCII-only source: non-ASCII characters
# are \u escapes.

#' Built-in string table for English
#' @noRd
.strings_en <- function() {
  list(
    ask_if = "Ask if",
    constraint = "Constraint",
    choice_filter = "Choices shown if",
    includes = "includes",
    does_not_include = "does not include",
    is = "is",
    is_not = "is not",
    is_empty = "is empty",
    is_not_empty = "is not empty",
    empty = "empty",
    the_answer = "the answer",
    this_answer = "this answer",
    number_selected_in = "number selected in",
    today = "today",
    now = "now",
    true = "true",
    false = "false",
    and = "and",
    or = "or",
    not = "not",
    repeat_for_each = "Repeat for each",
    instance_no = "Instance no.",
    other_specify = "Other (specify)",
    legend_required = "* Required question",
    legend_circle = "Choose one answer",
    legend_square = "Choose all answers that apply",
    legend_rank = paste("Rank the answers: write 1 in the box of the first",
                        "choice, 2 for the second, and so on"),
    legend_ask_if = "Ask if: ask the question only when the condition is met",
    not_on_paper = "Captured on the device, not on paper.",
    reference_no = "Reference no.",
    code = "Code",
    latitude = "Latitude",
    longitude = "Longitude",
    accuracy_m = "Accuracy (m)",
    range_steps = "%s to %s, steps of %s",
    range_from = "from",
    range_to = "to",
    at_least = "at least",
    at_most = "at most",
    more_than = "more than",
    less_than = "less than",
    fmt_date = "dd/mm/yyyy",
    fmt_month_year = "mm/yyyy",
    fmt_year = "yyyy",
    fmt_time = "hh:mm",
    choices_from = "Choices from",
    calculated = "Calculated",
    page = "Page",
    of = "of",
    version = "Version",
    generated = "Generated",
    form_id = "Form ID",
    language = "Language",
    image = "image"
  )
}

#' Built-in string table for French
#' @noRd
.strings_fr <- function() {
  list(
    ask_if = "Poser si",
    constraint = "Contrainte",
    choice_filter = "Choix propos\u00e9s si",
    includes = "comprend",
    does_not_include = "ne comprend pas",
    is = "est",
    is_not = "n'est pas",
    is_empty = "est vide",
    is_not_empty = "n'est pas vide",
    empty = "vide",
    the_answer = "la r\u00e9ponse",
    this_answer = "cette r\u00e9ponse",
    number_selected_in = "nombre de choix s\u00e9lectionn\u00e9s dans",
    today = "aujourd'hui",
    now = "maintenant",
    true = "vrai",
    false = "faux",
    and = "et",
    or = "ou",
    not = "non",
    repeat_for_each = "R\u00e9p\u00e9ter pour chaque",
    instance_no = "Occurrence n\u00b0",
    other_specify = "Autre (pr\u00e9ciser)",
    legend_required = "* Question obligatoire",
    legend_circle = "Choisir une seule r\u00e9ponse",
    legend_square = "Cocher toutes les r\u00e9ponses applicables",
    legend_rank = paste("Classer les r\u00e9ponses : \u00e9crire 1 dans la",
                        "case du premier choix, 2 pour le deuxi\u00e8me,",
                        "et ainsi de suite"),
    legend_ask_if = paste("Poser si : ne poser la question que si la",
                          "condition est remplie"),
    not_on_paper = "Saisi sur l'appareil, pas sur papier.",
    reference_no = "R\u00e9f\u00e9rence n\u00b0",
    code = "Code",
    latitude = "Latitude",
    longitude = "Longitude",
    accuracy_m = "Pr\u00e9cision (m)",
    range_steps = "de %s \u00e0 %s, par pas de %s",
    range_from = "de",
    range_to = "\u00e0",
    at_least = "au moins",
    at_most = "au plus",
    more_than = "plus de",
    less_than = "moins de",
    fmt_date = "jj/mm/aaaa",
    fmt_month_year = "mm/aaaa",
    fmt_year = "aaaa",
    fmt_time = "hh:mm",
    choices_from = "Choix provenant de",
    calculated = "Calcul\u00e9",
    page = "Page",
    of = "sur",
    version = "Version",
    generated = "G\u00e9n\u00e9r\u00e9 le",
    form_id = "Identifiant du formulaire",
    language = "Langue",
    image = "image"
  )
}

#' Names of the strings that can be overridden
#' @noRd
.strings_names <- function() {
  names(.strings_en())
}

#' The string table for a language
#'
#' @param lang_code Language code such as `"fr"` (NA, NULL or an unknown code
#'   gives English). Only the primary subtag is used, so `"fr-CA"` is French.
#' @param overrides NULL or a named list of single strings that replace
#'   entries of the table. Unknown names are an error.
#' @return A named list of strings.
#' @noRd
.cxf_strings <- function(lang_code = NULL, overrides = NULL) {
  code <- if (is.null(lang_code) || length(lang_code) != 1L ||
              is.na(lang_code)) {
    "en"
  } else {
    sub("[-_].*$", "", tolower(trimws(lang_code)))
  }
  out <- switch(code, fr = .strings_fr(), .strings_en())
  if (!is.null(overrides)) {
    bad <- setdiff(names(overrides), names(out))
    if (length(bad) > 0L || is.null(names(overrides))) {
      valid_txt <- paste(names(out), collapse = ", ")
      .cxf_abort(c(
        "Unknown string name{?s} in {.arg strings}: {.val {bad}}.",
        i = "Valid names are: {valid_txt}."
      ))
    }
    for (nm in names(overrides)) {
      out[[nm]] <- overrides[[nm]]
    }
  }
  out
}

#' Language code of the form in use ("en" when unknown)
#'
#' Taken from the `code` column of `form$languages` for the chosen suffix. A
#' name alone (for example "French") is recognized for English and French.
#' @noRd
.form_lang_code <- function(form) {
  sfx <- form$lang$suffix
  langs <- form$languages
  if (is.null(sfx) || is.na(sfx) || is.null(langs) || nrow(langs) == 0L) {
    return("en")
  }
  i <- match(tolower(sfx), tolower(langs$suffix))
  if (is.na(i)) {
    return("en")
  }
  code <- langs$code[i]
  if (is.na(code)) {
    nm <- tolower(langs$name[i])
    if (is.na(nm)) {
      return("en")
    }
    return(switch(nm, french = , "fran\u00e7ais" =, francais = "fr", "en"))
  }
  sub("[-_].*$", "", tolower(code))
}
