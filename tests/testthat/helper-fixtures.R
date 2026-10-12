# Fixture builders. Each writes a small XLSForm to a temp file and returns its path.

fixture_survey <- function() {
  data.frame(
    type = c("begin_group", "text", "select_one [yesno]", "end_group"),
    name = c("grp1", "q1_name", "q2_name", "end_grp1"),
    label = c("Household", "What is your name?", "Consent obtained?", "ignored"),
    hint = c(NA, "Full legal name", NA, NA),
    required = c(NA, "yes", "true", NA),
    stringsAsFactors = FALSE
  )
}

fixture_choices <- function() {
  data.frame(
    list_name = c("yesno", "yesno"),
    name = c("yes", "no"),
    label = c("Yes", "No"),
    stringsAsFactors = FALSE
  )
}

fixture_settings <- function() {
  data.frame(form_title = "Sample Form", stringsAsFactors = FALSE)
}

write_fixture <- function(sheets) {
  skip_if_not_installed("writexl")
  path <- tempfile(fileext = ".xlsx")
  writexl::write_xlsx(sheets, path = path)
  path
}

# Full form: survey + choices + settings, plain `label` column
make_basic_xlsform <- function() {
  write_fixture(list(
    survey = fixture_survey(),
    choices = fixture_choices(),
    settings = fixture_settings()
  ))
}

# Survey with `label::en` only (no plain label)
make_en_xlsform <- function() {
  survey <- fixture_survey()
  names(survey)[names(survey) == "label"] <- "label::en"
  write_fixture(list(survey = survey, choices = fixture_choices()))
}

# Survey with `label::en` and `label::sw`; choices with the same
make_bilingual_xlsform <- function() {
  survey <- fixture_survey()
  names(survey)[names(survey) == "label"] <- "label::en"
  survey[["label::sw"]] <- c("Kaya", "Jina lako ni nani?", "Idhini imepatikana?", "ignored")
  choices <- fixture_choices()
  names(choices)[names(choices) == "label"] <- "label::en"
  choices[["label::sw"]] <- c("Ndio", "Hapana")
  write_fixture(list(survey = survey, choices = choices))
}

make_missing_survey_xlsform <- function() {
  write_fixture(list(choices = fixture_choices()))
}

make_missing_type_xlsform <- function() {
  survey <- fixture_survey()
  survey <- survey[, setdiff(names(survey), "type")]
  write_fixture(list(survey = survey, choices = fixture_choices()))
}

make_nolabel_xlsform <- function() {
  survey <- fixture_survey()
  survey$label <- NULL
  write_fixture(list(survey = survey, choices = fixture_choices()))
}

# Build a normalized survey data frame directly, for .build_structure() tests.
make_struct_survey <- function(type, name = NULL, label = NULL) {
  n <- length(type)
  tibble::tibble(
    type = type,
    name = name %||% paste0("n", seq_len(n)),
    label_selected = label %||% paste0("L", seq_len(n)),
    repeat_count = NA_character_,
    base_type = convertxlsform:::.parse_base_type(type),
    type_arg = convertxlsform:::.parse_type_arg(type)
  )
}
