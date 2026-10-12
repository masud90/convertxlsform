make_png <- function(w, h) {
  skip_if_not(capabilities("png"))
  f <- tempfile(fileext = ".png")
  grDevices::png(f, width = w, height = h)
  graphics::par(mar = c(0, 0, 0, 0))
  graphics::plot.new()
  grDevices::dev.off()
  f
}

make_jpeg <- function(w, h) {
  skip_if_not(capabilities("jpeg"))
  f <- tempfile(fileext = ".jpg")
  grDevices::jpeg(f, width = w, height = h)
  graphics::par(mar = c(0, 0, 0, 0))
  graphics::plot.new()
  grDevices::dev.off()
  f
}

test_that("PNG size is read from the header", {
  f <- make_png(120, 80)
  expect_equal(.image_size_px(f), c(width = 120, height = 80))
  f2 <- make_png(1000, 33)
  expect_equal(.image_size_px(f2), c(width = 1000, height = 33))
})

test_that("JPEG size is read from the frame marker", {
  f <- make_jpeg(150, 90)
  expect_equal(.image_size_px(f), c(width = 150, height = 90))
  f2 <- make_jpeg(333, 777)
  expect_equal(.image_size_px(f2), c(width = 333, height = 777))
})

test_that("shipped example images are read", {
  d <- system.file("extdata", package = "convertxlsform")
  skip_if(!nzchar(d))
  expect_equal(.image_size_px(file.path(d, "all_types", "tap.png")),
               c(width = 200, height = 150))
  expect_equal(.image_size_px(file.path(d, "multilang", "apple.png")),
               c(width = 384, height = 131))
  expect_equal(.image_size_px(file.path(d, "all_types", "map.jpg")),
               c(width = 200, height = 150))
  # progressive JPEG with EXIF
  expect_equal(.image_size_px(file.path(d, "multilang", "survey.jpg")),
               c(width = 3456, height = 2304))
})

test_that("other formats, bad files and missing files give NA", {
  na <- c(width = NA_real_, height = NA_real_)
  txt <- tempfile(fileext = ".gif")
  writeLines("GIF89a not really an image", txt)
  expect_equal(.image_size_px(txt), na)
  empty <- tempfile()
  file.create(empty)
  expect_equal(.image_size_px(empty), na)
  expect_equal(.image_size_px(file.path(tempdir(), "nope.png")), na)
  expect_equal(.image_size_px(tempdir()), na)
  expect_equal(.image_size_px(NA_character_), na)
  # truncated PNG: signature only
  trunc <- tempfile(fileext = ".png")
  writeBin(as.raw(c(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0)), trunc)
  expect_equal(.image_size_px(trunc), na)
  # truncated JPEG: SOI then a segment that ends early
  tj <- tempfile(fileext = ".jpg")
  writeBin(as.raw(c(0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A)), tj)
  expect_equal(.image_size_px(tj), na)
})

# A form object is enough for .media_record()
fake_form <- function(root) list(media_root = root)

test_that("media records are found by name and by basename", {
  d <- tempfile("media")
  dir.create(d)
  f <- make_png(100, 50)
  file.copy(f, file.path(d, "pic.png"))
  form <- fake_form(d)
  rec <- .media_record(form, "pic.png", 2, 6)
  expect_true(rec$found)
  expect_equal(rec$file, "pic.png")
  expect_equal(basename(rec$path), "pic.png")
  expect_equal(rec$width_in, 100 / 96)
  expect_equal(rec$height_in, 50 / 96)
  # as given with a folder that does not exist: basename under media_root
  rec2 <- .media_record(form, "media/other/pic.png", 2, 6)
  expect_true(rec2$found)
  expect_equal(rec2$file, "media/other/pic.png")
  # a subfolder as given
  dir.create(file.path(d, "sub"))
  file.copy(f, file.path(d, "sub", "deep.png"))
  expect_true(.media_record(form, "sub/deep.png", 2, 6)$found)
  # blank means no image
  expect_null(.media_record(form, NA_character_, 2, 6))
  expect_null(.media_record(form, "", 2, 6))
})

test_that("sizes use 96 dpi and scale down to fit both limits, never up", {
  d <- tempfile("media")
  dir.create(d)
  # 960 x 96 px = 10 x 1 in; 96 x 960 px = 1 x 10 in; 48 x 48 px = 0.5 in
  file.copy(make_png(960, 96), file.path(d, "wide.png"))
  file.copy(make_png(96, 960), file.path(d, "tall.png"))
  file.copy(make_png(48, 48), file.path(d, "small.png"))
  form <- fake_form(d)

  w <- .media_record(form, "wide.png", max_height_in = 2, max_width_in = 5)
  expect_equal(c(w$width_in, w$height_in), c(5, 0.5))
  t <- .media_record(form, "tall.png", max_height_in = 2, max_width_in = 5)
  expect_equal(c(t$width_in, t$height_in), c(0.2, 2))
  s <- .media_record(form, "small.png", max_height_in = 2, max_width_in = 5)
  expect_equal(c(s$width_in, s$height_in), c(0.5, 0.5))
  # aspect ratio is kept
  expect_equal(w$width_in / w$height_in, 10)
  # choices use the choice image limit (0.6 in)
  c1 <- .media_record(form, "tall.png",
                      max_height_in = .cxf_choice_image_max_in,
                      max_width_in = 5)
  expect_equal(.cxf_choice_image_max_in, 0.6)
  expect_equal(c(c1$width_in, c1$height_in), c(0.06, 0.6))
})

test_that("a file in an unreadable format is found but has no size", {
  d <- tempfile("media")
  dir.create(d)
  writeLines("GIF89a", file.path(d, "a.gif"))
  rec <- .media_record(fake_form(d), "a.gif", 2, 6)
  expect_true(rec$found)
  expect_true(is.na(rec$width_in))
  expect_true(is.na(rec$height_in))
})

test_that("a missing file gives found = FALSE and is collected once", {
  d <- tempfile("media")
  dir.create(d)
  .media_reset_missing()
  rec <- .media_record(fake_form(d), "gone.png", 2, 6)
  expect_false(rec$found)
  expect_true(is.na(rec$path))
  expect_equal(rec$file, "gone.png")
  expect_true(is.na(rec$width_in))
  .media_record(fake_form(d), "gone.png", 2, 6)
  .media_record(fake_form(d), "also_gone.jpg", 2, 6)
  expect_equal(unique(.media_missing()), c("gone.png", "also_gone.jpg"))
  # one combined warning, then the list is reset
  w <- testthat::capture_warnings(.media_warn_missing())
  expect_length(w, 1L)
  expect_match(w, "2 media files not found")
  expect_match(w, "gone.png")
  expect_match(w, "also_gone.jpg")
  expect_length(.media_missing(), 0L)
  expect_silent(.media_warn_missing())
})

test_that("model building warns once for all missing media", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c("note", "select_one yn", "text", "select_one yn"),
    name = c("n", "q1", "q2", "q3"),
    label = c("Hello", "One", "Two", "Three"),
    image = c("missing_a.png", "missing_a.png", "missing_b.png", NA)
  )
  choices <- fdf(list_name = c("yn", "yn"), name = c("y", "n"),
                 label = c("Yes", "No"),
                 image = c("missing_c.png", "missing_a.png"))
  form <- read_quiet(make_form(survey, choices))
  warnings <- list()
  model <- withCallingHandlers(
    .build_model(form, test_opts()),
    warning = function(w) {
      warnings[[length(warnings) + 1L]] <<- conditionMessage(w)
      invokeRestart("muffleWarning")
    }
  )
  expect_length(warnings, 1L)
  expect_match(warnings[[1L]], "3 media files not found")
  for (f in c("missing_a.png", "missing_b.png", "missing_c.png")) {
    expect_match(warnings[[1L]], f, fixed = TRUE)
  }
  # the records are still in the model
  q1 <- block_named(model, "q1")
  expect_false(q1$image$found)
  expect_false(q1$answer$choices$image[[1L]]$found)
  expect_null(block_named(model, "q3")$image)
})

test_that("a missing choice file joins the single missing-files warning", {
  skip_if_no_writexl()
  survey <- fdf(type = c("select_one_from_file gone.csv", "text"),
                name = c("q1", "q2"), label = c("One", "Two"),
                image = c(NA, "gone.png"))
  form <- read_quiet(make_form(survey, fdf(list_name = "x", name = "a",
                                          label = "A")))
  warnings <- list()
  withCallingHandlers(
    .build_model(form, test_opts()),
    warning = function(w) {
      warnings[[length(warnings) + 1L]] <<- conditionMessage(w)
      invokeRestart("muffleWarning")
    }
  )
  expect_length(warnings, 1L)
  expect_match(warnings[[1L]], "gone.csv", fixed = TRUE)
  expect_match(warnings[[1L]], "gone.png", fixed = TRUE)
  expect_match(warnings[[1L]], "2 media files not found")
})
