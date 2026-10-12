test_that("symbols are the documented Unicode characters", {
  sym <- .cxf_symbols()
  expect_named(sym, c("circle", "square", "rank", "tick"))
  expect_equal(utf8ToInt(sym$circle), 0x25CB)
  expect_equal(utf8ToInt(sym$square), 0x25A1)
  expect_equal(utf8ToInt(sym$rank), 0x25A1)
  expect_equal(utf8ToInt(sym$tick), 0x25A1)
  expect_true(all(vapply(sym, length, integer(1)) == 1L))
})

test_that(".is_blank flags NA and whitespace-only strings", {
  expect_equal(.is_blank(c(NA, "", "   ", "\t", "a", " b ", NA_character_)),
               c(TRUE, TRUE, TRUE, TRUE, FALSE, FALSE, TRUE))
  expect_equal(.is_blank(NULL), logical(0))
  expect_equal(.is_blank(character(0)), logical(0))
  expect_equal(.is_blank(c(0, NA)), c(FALSE, TRUE))
})

test_that(".squish trims and collapses whitespace, NA-safe", {
  expect_equal(.squish("  a   b\t c  "), "a b c")
  expect_equal(.squish(c(" x ", NA, "y  z")), c("x", NA, "y z"))
  expect_equal(.squish(character(0)), character(0))
  expect_equal(.squish(NA_character_), NA_character_)
})

test_that(".excel_row maps data row i to Excel row i + 1", {
  expect_equal(.excel_row(1:3), c(2L, 3L, 4L))
  expect_type(.excel_row(1), "integer")
})

test_that("%||% from rlang is available and behaves as a default operator", {
  expect_equal(NULL %||% 2, 2)
  expect_equal(1 %||% 2, 1)
})

test_that(".cxf_inform is silent unless verbose", {
  expect_no_message(.cxf_inform(FALSE, "hidden"))
  expect_message(.cxf_inform(TRUE, "shown"), "shown")
})

test_that(".cxf_warn warns with the given text", {
  expect_warning(.cxf_warn("careful"), "careful")
})

test_that(".cxf_abort errors and interpolates the caller's variables", {
  expect_error(.cxf_abort("boom"), "boom")
  f <- function() {
    value <- "ZZ"
    .cxf_abort("bad value {value}")
  }
  expect_error(f(), "bad value ZZ")
})
