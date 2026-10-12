test_that("positive read returns survey, choices, settings and parsed types", {
  path <- make_basic_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, lang = "en", verbose = FALSE)

  expect_named(res, c("survey", "choices", "settings", "lang_used", "lang_mode"))
  expect_equal(nrow(res$survey), 4L)
  expect_equal(nrow(res$choices), 2L)
  expect_false(is.null(res$settings))
  expect_equal(res$survey$base_type, c("begin_group", "text", "select_one", "end_group"))
  expect_equal(res$survey$type_arg[3], "yesno")
  expect_equal(res$survey$required, c(NA, TRUE, TRUE, NA))
  expect_equal(res$survey$label_selected[2], "What is your name?")
  expect_length(res$survey$parameters_parsed, 4L)
})

test_that("exported convertxlsform() runs end to end and returns a path invisibly", {
  path <- make_basic_xlsform()
  expect_invisible(out <- convertxlsform(path, verbose = FALSE))
  expect_match(out, "\\.docx$")
  expect_message(convertxlsform(path, verbose = TRUE), "Completed in")
})

test_that("requested language present is used", {
  path <- make_bilingual_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, lang = "sw", verbose = FALSE)
  expect_equal(res$lang_mode, "multilingual")
  expect_equal(res$lang_used, "sw")
  expect_equal(res$survey$label_selected[2], "Jina lako ni nani?")
  expect_equal(res$choices$label_selected, c("Ndio", "Hapana"))
})

test_that("label::en is used when the requested language is 'en'", {
  path <- make_en_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, lang = "en", verbose = FALSE)
  expect_equal(res$lang_mode, "multilingual")
  expect_equal(res$lang_used, "en")
})

test_that("missing requested language falls back to English", {
  path <- make_en_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, lang = "sw", verbose = FALSE)
  expect_equal(res$lang_mode, "fallback_en")
  expect_equal(res$lang_used, "en")
  expect_equal(res$survey$label_selected[2], "What is your name?")
  expect_message(
    convertxlsform:::.read_and_validate_xlsform(path, lang = "sw", verbose = TRUE),
    "Falling back to English"
  )
})

test_that("single-language `label` form is used whatever language is requested", {
  path <- make_basic_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, lang = "sw", verbose = FALSE)
  expect_equal(res$lang_mode, "single_label")
  expect_equal(res$lang_used, "plain")
  expect_equal(res$survey$label_selected[3], "Consent obtained?")
})

test_that("'English (en)' is normalized to 'en'", {
  expect_equal(convertxlsform:::.normalize_lang("English (en)"), "en")
  expect_equal(convertxlsform:::.normalize_lang("English"), "en")
  expect_equal(convertxlsform:::.normalize_lang(" SW "), "sw")
  expect_equal(convertxlsform:::.normalize_lang(NULL), "en")
})

test_that("missing file errors", {
  expect_error(convertxlsform("does_not_exist.xlsx", verbose = FALSE), "Input file not found")
  expect_error(convertxlsform(NULL, verbose = FALSE), "Input file path is NULL")
})

test_that("missing survey sheet errors", {
  path <- make_missing_survey_xlsform()
  expect_error(convertxlsform(path, verbose = FALSE), "Missing required sheet\\(s\\): survey")
})

test_that("missing choices sheet errors", {
  path <- write_fixture(list(survey = fixture_survey()))
  expect_error(convertxlsform(path, verbose = FALSE), "Missing required sheet\\(s\\): choices")
})

test_that("missing type column errors", {
  path <- make_missing_type_xlsform()
  expect_error(convertxlsform(path, verbose = FALSE), "Missing columns in sheet 'survey': type")
})

test_that("missing all label columns errors", {
  path <- make_nolabel_xlsform()
  expect_error(convertxlsform(path, verbose = FALSE), "Missing label column in sheet 'survey'")
})

test_that("missing choices label columns errors", {
  choices <- fixture_choices()
  choices$label <- NULL
  path <- write_fixture(list(survey = fixture_survey(), choices = choices))
  expect_error(convertxlsform(path, verbose = FALSE), "Missing label column in sheet 'choices'")
})

test_that("column names are normalized to lower case", {
  survey <- fixture_survey()
  names(survey) <- toupper(names(survey))
  path <- write_fixture(list(survey = survey, choices = fixture_choices()))
  res <- convertxlsform:::.read_and_validate_xlsform(path, verbose = FALSE)
  expect_true(all(c("type", "name", "label") %in% names(res$survey)))
})

test_that("optional columns are created when absent", {
  path <- make_basic_xlsform()
  res <- convertxlsform:::.read_and_validate_xlsform(path, verbose = FALSE)
  expect_true(all(c("relevant", "constraint", "choice_filter", "image") %in% names(res$survey)))
  expect_true(all(c("image", "audio", "video") %in% names(res$choices)))
})

test_that(".ensure_columns adds missing columns with the right type", {
  df <- data.frame(a = 1)
  out <- convertxlsform:::.ensure_columns(df, c("b"), "character")
  expect_type(out$b, "character")
  out <- convertxlsform:::.ensure_columns(df, c("b"), "numeric")
  expect_type(out$b, "double")
  out <- convertxlsform:::.ensure_columns(df, c("b"), "logical")
  expect_type(out$b, "logical")
  expect_null(convertxlsform:::.ensure_columns(NULL, "b"))
})

# ---- unit tests for parsers ----

test_that(".parse_base_type extracts and lower-cases the base type", {
  f <- convertxlsform:::.parse_base_type
  expect_equal(f(c("text", "select_one yesno", "select_multiple [x]", "Begin_Group")),
               c("text", "select_one", "select_multiple", "begin_group"))
  expect_equal(f("dateTime"), "datetime")
  expect_equal(f("  text"), "text")
  expect_true(is.na(f(NA_character_)))
})

test_that(".parse_type_arg extracts bracketed argument", {
  f <- convertxlsform:::.parse_type_arg
  expect_equal(f(c("select_one [yesno]", "select_multiple [a_b]")), c("yesno", "a_b"))
  expect_equal(f(c("text", "select_one [yesno]")), c(NA, "yesno"))
  expect_equal(f("text"), NA_character_)
})

test_that(".parse_parameters_vectorized parses key=value pairs", {
  f <- convertxlsform:::.parse_parameters_vectorized
  expect_equal(f("start=1 end=10 step=1"), list(start = 1, end = 10, step = 1))
  expect_equal(f("Start=1, End=10"), list(start = 1, end = 10))
  expect_equal(f("randomize=true seed=12"), list(randomize = "true", seed = 12))
  expect_equal(f(NA_character_), list())
  expect_equal(f(""), list())
  expect_equal(f("junk a=b=c"), list())
})

test_that(".coerce_logical_maybe handles truthy, falsy and unknown values", {
  f <- convertxlsform:::.coerce_logical_maybe
  expect_equal(f(c("yes", "TRUE", "1", "Y", "t")), rep(TRUE, 5))
  expect_equal(f(c("no", "False", "0", "n", "F")), rep(FALSE, 5))
  expect_equal(f(c("maybe", NA, "")), c(NA, NA, NA))
  expect_equal(f(c(TRUE, FALSE)), c(TRUE, FALSE))
  expect_length(f(NULL), 0L)
})

test_that(".format_duration formats h:mm:ss", {
  f <- convertxlsform:::.format_duration
  expect_equal(f(0), "0:00:00")
  expect_equal(f(59), "0:00:59")
  expect_equal(f(61), "0:01:01")
  expect_equal(f(3661), "1:01:01")
  expect_equal(f(as.difftime(125, units = "secs")), "0:02:05")
})

test_that("%||% returns the fallback only for NULL", {
  expect_equal(NULL %||% 1, 1)
  expect_equal(NA %||% 1, NA)
})
