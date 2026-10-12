# Print the outline in two blocks so that nothing is cut at the line width:
# the structure (labels shortened) and the logic text.
print_outline <- function(x) {
  withr::local_options(width = 200)
  x <- as.data.frame(x)
  struct <- x[, c("row", "serial", "depth", "block", "kind", "name", "label",
                  "required", "n_choices")]
  struct$label <- ifelse(nchar(struct$label) > 40,
                         paste0(substr(struct$label, 1, 37), "..."),
                         struct$label)
  print(struct, right = FALSE)
  logic <- x[!is.na(x$relevant_text) | !is.na(x$constraint_text),
             c("row", "serial", "relevant_text", "constraint_text")]
  cut <- function(v) ifelse(nchar(v) > 140, paste0(substr(v, 1, 137), "..."), v)
  logic$relevant_text <- cut(logic$relevant_text)
  logic$constraint_text <- cut(logic$constraint_text)
  print(logic, right = FALSE)
  invisible(x)
}

example_path <- function(dir, file) {
  p <- system.file("extdata", dir, file, package = "convertxlsform")
  skip_if(!nzchar(p), paste(dir, "example not installed"))
  p
}

test_that("outline of the all_types form in English", {
  p <- example_path("all_types", "all_types.xlsx")
  expect_snapshot(print_outline(xlsform_outline(p, lang = "en")))
})

test_that("outline of the all_types form in French", {
  skip_if_not(l10n_info()[["UTF-8"]])
  p <- example_path("all_types", "all_types.xlsx")
  expect_snapshot(print_outline(xlsform_outline(p, lang = "fr")))
})

test_that("outline of the multilang form in English", {
  p <- example_path("multilang", "multilang.xlsx")
  expect_snapshot(print_outline(xlsform_outline(p, lang = "en")))
})

test_that("outline of the multilang form in Bengali", {
  skip_if_not(l10n_info()[["UTF-8"]])
  p <- example_path("multilang", "multilang.xlsx")
  expect_snapshot(print_outline(xlsform_outline(p, lang = "bn")))
})

test_that("outline returns a data frame with the documented columns", {
  p <- example_path("all_types", "all_types.xlsx")
  o <- xlsform_outline(p)
  expect_s3_class(o, "data.frame")
  expect_named(o, c("row", "serial", "depth", "block", "kind", "name", "label",
                    "required", "n_choices", "relevant_text",
                    "constraint_text", "layout"))
  expect_type(o$row, "integer")
  expect_type(o$serial, "character")
  expect_type(o$required, "logical")
  expect_type(o$n_choices, "integer")
  expect_true(all(o$block %in% c("section", "section_end", "note",
                                 "question")))
  expect_equal(o$n_choices[o$name == "main_crop" & o$block == "question"], 30L)
  expect_true(is.na(o$required[o$block == "section"][1]))
})

test_that("outline passes options through", {
  p <- example_path("all_types", "all_types.xlsx")
  flat <- xlsform_outline(p, nested_serial_numbers = FALSE)
  expect_true(all(is.na(flat$serial[flat$block == "section"])))
  q <- flat$serial[flat$block == "question"]
  expect_equal(q, as.character(seq_along(q)))

  nolog <- xlsform_outline(p, include_relevants = FALSE,
                           include_constraints = FALSE)
  expect_true(all(is.na(nolog$relevant_text)))
  expect_true(all(is.na(nolog$constraint_text)))

  hid <- xlsform_outline(p, include_hidden = TRUE)
  expect_gt(nrow(hid), nrow(xlsform_outline(p)))
  expect_true("hidden" %in% hid$kind)

  tr <- xlsform_outline(p, max_choices = 3)
  expect_equal(tr$n_choices[tr$name == "main_crop" & tr$block == "question"],
               30L)
})

test_that("outline reports problems against xlsform_outline()", {
  p <- example_path("all_types", "all_types.xlsx")
  e <- tryCatch(xlsform_outline(p, include_hints = "yes"),
                error = function(e) e)
  expect_s3_class(e, "rlang_error")
  expect_match(conditionMessage(e), "include_hints")
  expect_equal(rlang::call_name(e$call), "xlsform_outline")

  e <- tryCatch(xlsform_outline(p, bogus = 1), error = function(e) e)
  expect_equal(rlang::call_name(e$call), "xlsform_outline")
  expect_match(conditionMessage(e), "bogus")

  e <- tryCatch(xlsform_outline("no_such_file.xlsx"), error = function(e) e)
  expect_equal(rlang::call_name(e$call), "xlsform_outline")
  expect_match(conditionMessage(e), "not found")
})

test_that("outline is quiet by default and warns about an unknown language", {
  p <- example_path("all_types", "all_types.xlsx")
  expect_no_message(xlsform_outline(p))
  expect_warning(xlsform_outline(p, lang = "klingon"), "not found")
})
