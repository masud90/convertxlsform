test_that("multilang example reads in English and Bengali", {
  p <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  langs <- xlsform_languages(p)
  expect_equal(langs$code, c("en", "bn"))
  expect_equal(langs$default, c(TRUE, FALSE))

  en <- read_quiet(p, lang = "en")
  bn <- read_quiet(p, lang = "Bengali")
  expect_equal(en$lang$suffix, "English (en)")
  expect_equal(bn$lang$suffix, "Bengali (bn)")
  expect_equal(nrow(en$survey), 46)
  expect_equal(nrow(bn$survey), 46)
  expect_equal(nrow(en$choices), 33)
  expect_false(identical(en$survey$label, bn$survey$label))
  expect_true(any(!is.na(en$survey$serial)))
  expect_equal(sum(!is.na(en$survey$serial)), sum(!is.na(bn$survey$serial)))
  expect_equal(en$survey$serial[which(en$survey$name == "groupA")], "1")
  expect_equal(en$survey$serial[which(en$survey$name == "mem_name")], "2.2.1")
  expect_equal(en$settings$form_title,
               "ConvertXLSForm Tool Testing Questionnaire")
  expect_equal(en$settings$default_language, "English (en)")
})

test_that("all_types example reads when present", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p), "all_types example not installed")
  f <- read_quiet(p)
  expect_s3_class(f, "cxf_form")
  expect_gt(nrow(f$survey), 20)
  expect_false("unknown" %in% f$survey$kind)
  expect_true(any(!is.na(f$survey$serial)))
})
