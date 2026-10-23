# Fonts for complex scripts: which font to ask for, whether it is installed,
# and which font files to hand to LibreOffice. ASCII-only source.

#' Complex scripts recognized in labels
#'
#' Each script has its Unicode block, its direction and a list of candidate
#' fonts, best first: the Noto font, then fonts that ship with macOS and
#' Windows, then common Linux fonts.
#' @noRd
.cs_scripts <- function() {
  list(
    bengali = list(lo = 0x0980L, hi = 0x09FFL, rtl = FALSE,
                   fonts = c("Noto Sans Bengali", "Noto Serif Bengali",
                             "Kohinoor Bangla", "Bangla Sangam MN",
                             "Bangla MN", "Nirmala UI", "Vrinda",
                             "Lohit Bengali", "FreeSans")),
    arabic = list(lo = 0x0600L, hi = 0x06FFL, rtl = TRUE,
                  fonts = c("Noto Naskh Arabic", "Noto Sans Arabic",
                            "Geeza Pro", "Arial", "Tahoma",
                            "Times New Roman", "DejaVu Sans")),
    devanagari = list(lo = 0x0900L, hi = 0x097FL, rtl = FALSE,
                      fonts = c("Noto Sans Devanagari",
                                "Noto Serif Devanagari",
                                "Kohinoor Devanagari",
                                "Devanagari Sangam MN", "Nirmala UI",
                                "Mangal", "Lohit Devanagari", "FreeSans")),
    ethiopic = list(lo = 0x1200L, hi = 0x137FL, rtl = FALSE,
                    fonts = c("Noto Sans Ethiopic", "Noto Serif Ethiopic",
                              "Kefa", "Ebrima", "Nyala", "Abyssinica SIL",
                              "FreeSerif")),
    hebrew = list(lo = 0x0590L, hi = 0x05FFL, rtl = TRUE,
                  fonts = c(NA_character_, "Noto Sans Hebrew", "Arial",
                            "DejaVu Sans"))
  )
}

# Cache of the installed font table (filled on first use).
.cxf_font_cache <- new.env(parent = emptyenv())

#' Installed fonts: a data.frame with `family` and `path`
#'
#' Read once per session with systemfonts. An empty table when the fonts
#' cannot be listed.
#' @noRd
.installed_fonts <- function() {
  if (is.null(.cxf_font_cache$table)) {
    tab <- tryCatch(systemfonts::system_fonts(), error = function(e) NULL)
    .cxf_font_cache$table <- if (is.null(tab) || nrow(tab) == 0L) {
      data.frame(family = character(0), path = character(0),
                 stringsAsFactors = FALSE)
    } else {
      data.frame(family = as.character(tab$family),
                 path = as.character(tab$path), stringsAsFactors = FALSE)
    }
  }
  .cxf_font_cache$table
}

#' Is a font family installed?
#' @param family Character vector of family names.
#' @return Logical vector.
#' @noRd
.font_installed <- function(family) {
  fams <- tolower(.installed_fonts()$family)
  !is.na(family) & tolower(family) %in% fams
}

#' Font files of the given families (all styles)
#' @noRd
.font_files <- function(family) {
  tab <- .installed_fonts()
  keep <- tolower(tab$family) %in% tolower(family[!is.na(family)])
  files <- unique(tab$path[keep])
  files[file.exists(files)]
}

#' Script-dependent font for complex-script text
#'
#' The script with the most characters in `text` decides. Its first installed
#' candidate font is used; when none is installed, the first candidate (the
#' Noto font) is named so the document asks for the font to install.
#'
#' @param text Character vector of labels in the language in use.
#' @param font_family The main font family (used for Hebrew and as the
#'   result when no complex script is found).
#' @param installed Function telling which fonts are installed; tests pass
#'   their own.
#' @return list(font_family_cs, rtl, script, installed); `script` is NA when
#'   no complex script is found, and `installed` says whether the chosen font
#'   was found on this machine.
#' @noRd
.detect_script_font <- function(text, font_family,
                                installed = .font_installed) {
  text <- text[!is.na(text)]
  cp <- if (length(text) > 0L) {
    unlist(lapply(enc2utf8(text), utf8ToInt), use.names = FALSE)
  } else {
    integer(0)
  }
  scripts <- .cs_scripts()
  counts <- vapply(scripts, function(z) sum(cp >= z$lo & cp <= z$hi), 1L)
  if (all(counts == 0L)) {
    return(list(font_family_cs = font_family, rtl = FALSE,
                script = NA_character_, installed = TRUE))
  }
  key <- names(scripts)[which.max(counts)]
  z <- scripts[[key]]
  cands <- z$fonts
  cands[is.na(cands)] <- font_family
  cands <- unique(cands)
  ok <- installed(cands)
  font <- if (any(ok)) cands[which(ok)[1L]] else cands[1L]
  list(font_family_cs = font, rtl = z$rtl, script = key,
       installed = any(ok))
}

#' Warn when the complex-script font of a PDF is missing
#'
#' @param model A `cxf_model`.
#' @return The files of the complex-script font, to give LibreOffice,
#'   invisibly.
#' @noRd
.check_pdf_fonts <- function(model) {
  m <- model$meta
  script <- m$script %||% NA_character_
  if (is.na(script)) {
    return(invisible(character(0)))
  }
  files <- .font_files(m$font_family_cs)
  if (length(files) == 0L) {
    font <- m$font_family_cs
    .cxf_warn(c(
      "The font {.val {font}} for {script} text is not installed.",
      i = "The PDF may show blank text or boxes for that script.",
      i = paste("Install the font (for example from",
                "{.url https://fonts.google.com/noto}), or set",
                "{.arg font_family_cs} to an installed font.")
    ))
  }
  invisible(files)
}
