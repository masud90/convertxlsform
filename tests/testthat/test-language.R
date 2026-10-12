ml <- function(...) {
  fdf(type = c("text", "select_one l"), name = c("a", "b"), ...)
}
ch2 <- fdf(list_name = c("l", "l"), name = c("1", "2"))

test_that("suffix parsing gives code and name", {
  p <- .parse_lang_suffix(c("English (en)", "en", "English", "pt-BR",
                            "Bengali (bn)"))
  expect_equal(p$code, c("en", "en", NA, "pt-BR", "bn"))
  expect_equal(p$name, c("English", NA, "English", NA, "Bengali"))
})

test_that("headers are normalized", {
  expect_equal(
    .norm_header(c(" Label::English (en) ", "LABEL", "media::image::en",
                   "Media::Image", "list name", "constraint_message::FR", NA)),
    c("label::English (en)", "label", "image::en", "image", "list_name",
      "constraint_message::FR", NA)
  )
})

test_that("language matching by suffix, code and name", {
  skip_if_no_writexl()
  sv <- ml(`label::English (en)` = c("A", "B"),
           `label::Bengali (bn)` = c("AA", "BB"))
  ch <- cbind(ch2, `label::English (en)` = c("Yes", "No"),
              `label::Bengali (bn)` = c("Ha", "Na"))
  p <- make_form(sv, ch)
  for (req in c("Bengali (bn)", "bn", "Bengali", "BN", "bengali")) {
    f <- read_quiet(p, lang = req)
    expect_equal(f$lang$suffix, "Bengali (bn)", info = req)
    expect_equal(f$lang$mode, "multilingual")
    expect_equal(f$survey$label, c("AA", "BB"))
    expect_equal(f$choices$label, c("Ha", "Na"))
  }
  f <- read_quiet(p)
  expect_equal(f$lang$suffix, "English (en)")
  expect_equal(f$survey$label, c("A", "B"))
  expect_equal(f$languages$code, c("en", "bn"))
})

test_that("label::en style headers and bare names work", {
  skip_if_no_writexl()
  p <- make_form(ml(`label::en` = c("A", "B"), `label::fr` = c("AF", "BF")),
                 cbind(ch2, `label::en` = c("Y", "N"), `label::fr` = c("O", "N")))
  f <- read_quiet(p, lang = "fr")
  expect_equal(f$survey$label, c("AF", "BF"))
  expect_equal(f$languages$suffix, c("en", "fr"))
  expect_true(is.na(f$languages$name[1]))
  p2 <- make_form(ml(`label::English` = c("A", "B")),
                  cbind(ch2, `label::English` = c("Y", "N")))
  f2 <- read_quiet(p2, lang = "english")
  expect_equal(f2$survey$label, c("A", "B"))
  expect_equal(f2$languages$name, "English")
})

test_that("a plain label column is single-language", {
  skip_if_no_writexl()
  p <- make_form(ml(label = c("A", "B")), cbind(ch2, label = c("Y", "N")))
  f <- read_quiet(p)
  expect_equal(f$lang$mode, "single")
  expect_equal(nrow(f$languages), 0)
  expect_equal(f$survey$label, c("A", "B"))
  expect_warning(f <- read_quiet(p, lang = "fr"), "no language-specific")
  expect_equal(f$survey$label, c("A", "B"))
})

test_that("an unknown language warns and falls back", {
  skip_if_no_writexl()
  p <- make_form(ml(`label::English (en)` = c("A", "B"),
                    `label::Bengali (bn)` = c("AA", "BB")),
                 cbind(ch2, `label::English (en)` = c("Y", "N"),
                       `label::Bengali (bn)` = c("H", "N")),
                 settings = fdf(default_language = "Bengali (bn)"))
  expect_warning(f <- read_quiet(p, lang = "zz"), "not found")
  expect_equal(f$lang$mode, "fallback")
  expect_equal(f$lang$requested, "zz")
  expect_equal(f$lang$suffix, "Bengali (bn)")
  expect_equal(f$survey$label, c("AA", "BB"))
})

test_that("blank translations fall back per cell", {
  skip_if_no_writexl()
  p <- make_form(
    ml(label = c("Plain A", "Plain B"),
       `label::English (en)` = c("A", "B"),
       `label::Bengali (bn)` = c("AA", NA),
       `hint::English (en)` = c("hint A", NA),
       `hint::Bengali (bn)` = c(NA, NA)),
    cbind(ch2, `label::English (en)` = c("Y", "N"),
          `label::Bengali (bn)` = c("H", NA))
  )
  f <- read_quiet(p, lang = "bn")
  expect_equal(f$survey$label, c("AA", "Plain B"))
  expect_equal(f$survey$hint, c("hint A", NA))
  expect_equal(f$choices$label, c("H", "N"))

  # No plain column: fall back to the default language
  p2 <- make_form(
    ml(`label::English (en)` = c("A", "B"), `label::Bengali (bn)` = c("AA", NA)),
    cbind(ch2, `label::English (en)` = c("Y", "N"),
          `label::Bengali (bn)` = c("H", NA)),
    settings = fdf(default_language = "English (en)")
  )
  f2 <- read_quiet(p2, lang = "bn")
  expect_equal(f2$survey$label, c("AA", "B"))
})

test_that("default_language from settings picks the language", {
  skip_if_no_writexl()
  sv <- ml(`label::English (en)` = c("A", "B"), `label::French (fr)` = c("AF", "BF"))
  ch <- cbind(ch2, `label::English (en)` = c("Y", "N"),
              `label::French (fr)` = c("O", "N"))
  f <- read_quiet(make_form(sv, ch, settings = fdf(default_language = "fr")))
  expect_equal(f$lang$suffix, "French (fr)")
  expect_equal(f$survey$label, c("AF", "BF"))
  # default_language that matches nothing: first language
  f <- read_quiet(make_form(sv, ch, settings = fdf(default_language = "de")))
  expect_equal(f$lang$suffix, "English (en)")
})

test_that("form_title::lang is honoured with fallbacks", {
  skip_if_no_writexl()
  sv <- ml(`label::English (en)` = c("A", "B"), `label::French (fr)` = c("AF", "BF"))
  ch <- cbind(ch2, `label::English (en)` = c("Y", "N"),
              `label::French (fr)` = c("O", "N"))
  st <- fdf(form_title = "Plain", `form_title::French (fr)` = "Titre",
            default_language = "English (en)")
  p <- make_form(sv, ch, settings = st)
  expect_equal(read_quiet(p, lang = "fr")$settings$form_title, "Titre")
  expect_equal(read_quiet(p, lang = "en")$settings$form_title, "Plain")
})

test_that("media::image columns are treated as image", {
  skip_if_no_writexl()
  p <- make_form(ml(`label::en` = c("A", "B"),
                    `media::image::en` = c("a.png", NA),
                    `media::audio` = c(NA, "b.mp3")),
                 cbind(ch2, `label::en` = c("Y", "N"),
                       `media::image` = c("y.png", NA)))
  f <- read_quiet(p)
  expect_equal(f$survey$image, c("a.png", NA))
  expect_equal(f$survey$audio, c(NA, "b.mp3"))
  expect_equal(f$choices$image, c("y.png", NA))
})

test_that("xlsform_languages lists languages and the default", {
  skip_if_no_writexl()
  sv <- ml(`label::English (en)` = c("A", "B"), `label::French (fr)` = c("AF", "BF"))
  ch <- cbind(ch2, `label::English (en)` = c("Y", "N"),
              `label::French (fr)` = c("O", "N"))
  p <- make_form(sv, ch, settings = fdf(default_language = "French (fr)"))
  l <- xlsform_languages(p)
  expect_equal(names(l), c("suffix", "code", "name", "default"))
  expect_equal(l$default, c(FALSE, TRUE))
  l1 <- xlsform_languages(make_form(sv, ch))
  expect_equal(l1$default, c(TRUE, FALSE))
  single <- xlsform_languages(make_form(ml(label = c("A", "B")),
                                        cbind(ch2, label = c("Y", "N"))))
  expect_equal(nrow(single), 0)
  expect_type(single$default, "logical")
  expect_equal(xlsform_languages(read_quiet(p)), l)
})

test_that("name-only headers get an ISO code so lang = 'en' matches", {
  skip_if_no_writexl()
  sv <- ml(`label::English` = c("A", "B"), `label::French` = c("AF", "BF"),
           `label::Bengali` = c("AB", "BB"))
  ch <- cbind(ch2, `label::English` = c("Y", "N"),
              `label::French` = c("O", "N"), `label::Bengali` = c("H", "N"))
  p <- make_form(sv, ch)
  f <- read_quiet(p, lang = "fr")
  expect_equal(f$languages$code, c("en", "fr", "bn"))
  expect_equal(f$lang$suffix, "French")
  expect_equal(read_quiet(p, lang = "en")$lang$suffix, "English")
  expect_equal(read_quiet(p, lang = "bn")$survey$label, c("AB", "BB"))
  # unknown names stay without a code
  d <- .detect_languages(c("label::Klingon", "label::en"))
  expect_equal(d$code, c(NA, "en"))
})
