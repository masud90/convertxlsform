mk <- function(kind, label = "L") {
  data.frame(row = seq_along(kind) + 1L, kind = kind,
             label = ifelse(kind %in% c("end_group", "end_repeat"), NA, label),
             stringsAsFactors = FALSE)
}

test_that("nested numbering follows the spec example", {
  s <- mk(c("text", "begin_group", "text", "begin_group", "text", "end_group",
            "end_group", "text"))
  out <- .build_structure(s, TRUE)
  expect_equal(out$serial,
               c("1", "2", "2.1", "2.2", "2.2.1", NA, NA, "3"))
  expect_equal(out$depth, c(0, 0, 1, 1, 2, 1, 0, 0))
  expect_equal(out$parent, c(NA, NA, 2, 2, 4, 2, NA, NA))
})

test_that("notes, hidden rows and end rows are unnumbered", {
  out <- .build_structure(mk(c("note", "hidden", "text", "unknown",
                               "begin_repeat", "geo", "end_repeat")), TRUE)
  expect_equal(out$serial, c(NA, NA, "1", "2", "3", "3.1", NA))
})

test_that("unlabelled groups are transparent", {
  s <- mk(c("begin_group", "text", "text", "end_group", "text"))
  s$label[1] <- NA
  out <- .build_structure(s, TRUE)
  expect_equal(out$serial, c(NA, "1", "2", NA, "3"))
  expect_equal(out$depth, c(0, 1, 1, 0, 0))

  s2 <- mk(c("begin_group", "begin_group", "text", "end_group", "end_group"))
  s2$label[2] <- NA
  expect_equal(.build_structure(s2, TRUE)$serial, c("1", NA, "1.1", NA, NA))
})

test_that("flat numbering numbers questions only", {
  s <- mk(c("text", "begin_group", "text", "begin_repeat", "text",
            "end_repeat", "end_group", "text"))
  out <- .build_structure(s, FALSE)
  expect_equal(out$serial, c("1", NA, "2", NA, "3", NA, NA, "4"))
})

test_that("structural errors cite Excel rows", {
  expect_error(.build_structure(mk(c("text", "end_group")), TRUE),
               "Row 3")
  expect_error(.build_structure(mk(c("begin_group", "text")), TRUE),
               "Row 2.*never closed")
  expect_error(.build_structure(mk(c("begin_group", "end_repeat")), TRUE),
               "Row 3.*end_repeat.*begin_group.*row 2")
  expect_error(.build_structure(mk(c("begin_repeat", "end_group")), TRUE),
               "end_group")
  e <- tryCatch(.build_structure(mk(c("end_group", "begin_repeat")), TRUE),
                error = function(e) conditionMessage(e))
  expect_match(e, "Row 2")
  expect_match(e, "Row 3")
})

test_that("a 5,000 row form is numbered quickly", {
  # make it balanced: build whole blocks of 7
  kinds <- rep(c("begin_group", "text", "text", "begin_repeat", "text",
                 "end_repeat", "end_group"), 715)
  t <- system.time(out <- .build_structure(mk(kinds), TRUE))[["elapsed"]]
  expect_lt(t, 1)
  expect_equal(nrow(out), 5005)
  expect_equal(out$serial[1:5], c("1", "1.1", "1.2", "1.3", "1.3.1"))
})
