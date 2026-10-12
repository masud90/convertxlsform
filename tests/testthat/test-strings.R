test_that("English and French tables have the same names", {
  en <- .cxf_strings("en")
  fr <- .cxf_strings("fr")
  expect_type(en, "list")
  expect_equal(names(fr), names(en))
  expect_true(all(vapply(en, function(v) is.character(v) && length(v) == 1L,
                         NA)))
  expect_true(all(vapply(fr, function(v) is.character(v) && length(v) == 1L,
                         NA)))
  expect_equal(names(en), .strings_names())
  expect_equal(en$ask_if, "Ask if")
  expect_equal(fr$ask_if, "Poser si")
  expect_equal(fr$other_specify, "Autre (pr\u00e9ciser)")
  # the two tables differ where it matters
  expect_false(identical(en$includes, fr$includes))
  # every string needed by the spec exists
  needed <- c("ask_if", "constraint", "includes", "does_not_include", "is",
              "is_not", "is_empty", "is_not_empty", "the_answer",
              "number_selected_in", "today", "repeat_for_each", "instance_no",
              "other_specify", "not_on_paper", "reference_no", "code")
  expect_true(all(needed %in% names(en)))
  expect_true(any(startsWith(names(en), "legend")))
})

test_that("unknown languages fall back to English", {
  expect_equal(.cxf_strings("bn"), .cxf_strings("en"))
  expect_equal(.cxf_strings(NA), .cxf_strings("en"))
  expect_equal(.cxf_strings(NULL), .cxf_strings("en"))
  expect_equal(.cxf_strings("xx"), .cxf_strings("en"))
  # case and region subtags
  expect_equal(.cxf_strings("FR"), .cxf_strings("fr"))
  expect_equal(.cxf_strings("fr-CA"), .cxf_strings("fr"))
})

test_that("overrides replace entries and unknown names are an error", {
  s <- .cxf_strings("fr", list(ask_if = "Si", code = "Code postal"))
  expect_equal(s$ask_if, "Si")
  expect_equal(s$code, "Code postal")
  expect_equal(s$includes, .cxf_strings("fr")$includes)
  expect_equal(names(s), .strings_names())
  expect_error(.cxf_strings("en", list(bogus = "x")), "bogus")
  expect_error(.cxf_strings("en", list(bogus = "x")), "Valid names")
})

test_that("the language code comes from the form", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  expect_equal(.form_lang_code(read_quiet(p, lang = "fr")), "fr")
  expect_equal(.form_lang_code(read_quiet(p, lang = "en")), "en")
  q <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(q))
  expect_equal(.form_lang_code(read_quiet(q, lang = "bn")), "bn")
  skip_if_no_writexl()
  # single-language form: English
  expect_equal(.form_lang_code(small_form()), "en")
  # a name without a code
  survey <- fdf(type = "text", name = "q", `label::French` = "Question",
                `label::English` = "Question")
  f <- read_quiet(make_form(survey), lang = "French")
  expect_equal(.form_lang_code(f), "fr")
  g <- read_quiet(make_form(fdf(type = "text", name = "q",
                                `label::Klingon` = "Q")))
  expect_equal(.form_lang_code(g), "en")
})
