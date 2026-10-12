# PDF route: convert a .docx to PDF with LibreOffice, or Word via doconv.

#' Locate the LibreOffice binary
#'
#' Checks, in order: option `convertxlsform.soffice`, environment variable
#' `CONVERTXLSFORM_SOFFICE`, `soffice` and `libreoffice` on the PATH, then
#' the usual install locations for the platform.
#'
#' @return Path to the binary, or `NA_character_` when none is found.
#' @noRd
.find_soffice <- function() {
  usable <- function(p) {
    length(p) == 1L && !is.na(p) && nzchar(p) && file.exists(p)
  }

  opt <- getOption("convertxlsform.soffice")
  if (is.character(opt) && usable(opt[1])) {
    return(normalizePath(opt[1], winslash = "/", mustWork = FALSE))
  }

  env <- Sys.getenv("CONVERTXLSFORM_SOFFICE", unset = "")
  if (usable(env)) {
    return(normalizePath(env, winslash = "/", mustWork = FALSE))
  }

  for (nm in c("soffice", "libreoffice")) {
    p <- unname(Sys.which(nm))
    if (usable(p)) {
      return(p)
    }
  }

  sysname <- Sys.info()[["sysname"]]
  candidates <- character()
  if (identical(sysname, "Darwin")) {
    candidates <- c(
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      path.expand("~/Applications/LibreOffice.app/Contents/MacOS/soffice")
    )
  } else if (identical(sysname, "Windows")) {
    pf <- c(Sys.getenv("ProgramFiles"), Sys.getenv("ProgramFiles(x86)"))
    pf <- pf[nzchar(pf)]
    candidates <- c(
      file.path(pf, "LibreOffice", "program", "soffice.exe"),
      "C:/Program Files/LibreOffice/program/soffice.exe",
      "C:/Program Files (x86)/LibreOffice/program/soffice.exe"
    )
  } else {
    candidates <- c(
      "/usr/bin/soffice",
      "/usr/lib/libreoffice/program/soffice",
      Sys.glob("/opt/libreoffice*/program/soffice"),
      "/snap/bin/libreoffice"
    )
  }
  for (p in unique(candidates)) {
    if (usable(p)) {
      return(normalizePath(p, winslash = "/", mustWork = FALSE))
    }
  }
  NA_character_
}

#' Is LibreOffice available?
#'
#' Reports whether a LibreOffice binary can be found, which is needed to
#' convert Word documents to PDF with [docx_to_pdf()]. The search checks the
#' option `convertxlsform.soffice`, the environment variable
#' `CONVERTXLSFORM_SOFFICE`, the `PATH`, and the default install locations
#' for the operating system.
#'
#' @return `TRUE` or `FALSE`.
#' @export
#' @examples
#' has_libreoffice()
has_libreoffice <- function() {
  !is.na(.find_soffice())
}

#' Convert a Word document to PDF
#'
#' Converts a `.docx` file to PDF with LibreOffice, or with Microsoft Word
#' through the `doconv` package. LibreOffice is free and runs on all
#' platforms. Each call uses a fresh temporary LibreOffice profile, so it
#' works while a LibreOffice window is open and in parallel runs.
#'
#' If LibreOffice is installed in an unusual place, set the option
#' `convertxlsform.soffice` or the environment variable
#' `CONVERTXLSFORM_SOFFICE` to the path of the `soffice` binary.
#'
#' @param docx Path to an existing `.docx` file.
#' @param pdf Path of the PDF to write. `NULL` (default) uses the path of
#'   `docx` with the extension changed to `.pdf`.
#' @param engine `"auto"` (default) tries LibreOffice, then Word through
#'   `doconv` if that package is installed. `"libreoffice"` and `"word"`
#'   force one engine.
#' @param overwrite Overwrite an existing `pdf`? Default `FALSE`.
#' @param verbose Print progress messages? Default `TRUE`.
#'
#' @return The path of the PDF, invisibly.
#' @export
#' @examples
#' if (interactive() && has_libreoffice()) {
#'   docx <- tempfile(fileext = ".docx")
#'   doc <- officer::read_docx()
#'   doc <- officer::body_add_par(doc, "Hello")
#'   print(doc, target = docx)
#'   pdf <- docx_to_pdf(docx, pdf = tempfile(fileext = ".pdf"))
#'   file.exists(pdf)
#' }
docx_to_pdf <- function(docx,
                        pdf = NULL,
                        engine = c("auto", "libreoffice", "word"),
                        overwrite = FALSE,
                        verbose = TRUE) {
  engine <- match.arg(engine)

  if (!is.character(docx) || length(docx) != 1L || is.na(docx)) {
    .cxf_abort("{.arg docx} must be a single file path.")
  }
  if (!file.exists(docx) || dir.exists(docx)) {
    .cxf_abort("File {.file {docx}} does not exist.")
  }
  if (!grepl("\\.docx$", docx, ignore.case = TRUE)) {
    .cxf_abort("{.arg docx} must be a {.val .docx} file, not {.file {docx}}.")
  }
  docx <- normalizePath(docx, winslash = "/", mustWork = TRUE)

  if (is.null(pdf)) {
    pdf <- sub("\\.docx$", ".pdf", docx, ignore.case = TRUE)
  }
  if (!is.character(pdf) || length(pdf) != 1L || is.na(pdf)) {
    .cxf_abort("{.arg pdf} must be NULL or a single file path.")
  }
  if (file.exists(pdf) && !isTRUE(overwrite)) {
    .cxf_abort(c(
      "File {.file {pdf}} already exists.",
      "i" = "Use {.code overwrite = TRUE} to replace it."
    ))
  }
  if (!dir.exists(dirname(pdf))) {
    .cxf_abort("Folder {.file {dirname(pdf)}} does not exist.")
  }

  soffice <- if (engine %in% c("auto", "libreoffice")) .find_soffice() else NA_character_

  if (engine == "libreoffice" && is.na(soffice)) {
    .abort_no_engine(word = FALSE)
  }

  if (!is.na(soffice)) {
    .cxf_inform(verbose, "Converting to PDF with LibreOffice.")
    .soffice_convert(soffice, docx, pdf)
  } else {
    if (!requireNamespace("doconv", quietly = TRUE)) {
      .abort_no_engine(word = engine == "word")
    }
    .cxf_inform(verbose, "Converting to PDF with Word via {.pkg doconv}.")
    .word_convert(docx, pdf)
  }

  .cxf_inform(verbose, c("v" = "Wrote {.file {pdf}}."))
  invisible(pdf)
}

#' Run LibreOffice headless and copy the result to `pdf`
#' @noRd
.soffice_convert <- function(soffice, docx, pdf) {
  work <- tempfile("cxf_pdf_")
  profile <- tempfile("cxf_profile_")
  dir.create(work)
  dir.create(profile)
  on.exit(unlink(c(work, profile), recursive = TRUE, force = TRUE), add = TRUE)

  profile_url <- paste0(
    "file:///",
    sub("^/+", "", normalizePath(profile, winslash = "/", mustWork = FALSE))
  )
  args <- c(
    "--headless", "--norestore",
    shQuote(paste0("-env:UserInstallation=", profile_url)),
    "--convert-to", "pdf",
    "--outdir", shQuote(work),
    shQuote(docx)
  )
  out <- tryCatch(
    suppressWarnings(system2(soffice, args, stdout = TRUE, stderr = TRUE)),
    error = function(e) conditionMessage(e)
  )

  produced <- file.path(
    work,
    sub("\\.docx$", ".pdf", basename(docx), ignore.case = TRUE)
  )
  if (!file.exists(produced)) {
    captured <- paste(out, collapse = "\n")
    .cxf_abort(c(
      "LibreOffice did not produce a PDF.",
      "i" = "Binary: {.file {soffice}}",
      "i" = "Output: {captured}"
    ))
  }
  ok <- file.copy(produced, pdf, overwrite = TRUE)
  if (!ok) {
    .cxf_abort("Could not write {.file {pdf}}.")
  }
  invisible(pdf)
}

#' Convert with Word through doconv
#' @noRd
.word_convert <- function(docx, pdf) {
  res <- tryCatch(
    doconv::docx2pdf(input = docx, output = pdf),
    error = function(e) e
  )
  if (inherits(res, "error") || !file.exists(pdf)) {
    msg <- if (inherits(res, "error")) conditionMessage(res) else "no PDF produced"
    .cxf_abort(c(
      "Word conversion through {.pkg doconv} failed.",
      "i" = "{msg}"
    ))
  }
  invisible(pdf)
}

#' Abort with install hints when no conversion engine is available
#' @noRd
.abort_no_engine <- function(word = FALSE, call = rlang::caller_env()) {
  sysname <- Sys.info()[["sysname"]]
  hint <- if (identical(sysname, "Darwin")) {
    c(
      "i" = "macOS: {.code brew install --cask libreoffice}, or download from {.url https://www.libreoffice.org/download/}."
    )
  } else if (identical(sysname, "Windows")) {
    c("i" = "Windows: download from {.url https://www.libreoffice.org/download/}.")
  } else {
    c(
      "i" = "Linux: {.code sudo apt install libreoffice-writer}, or download from {.url https://www.libreoffice.org/download/}."
    )
  }
  head <- if (word) {
    "Word conversion needs the {.pkg doconv} package and Microsoft Word, and {.pkg doconv} is not installed."
  } else {
    "No LibreOffice installation was found, so the PDF cannot be created."
  }
  .cxf_abort(
    c(
      head,
      hint,
      "i" = "If LibreOffice is installed in an unusual place, set the option {.code convertxlsform.soffice} to the path of the {.file soffice} binary.",
      "i" = "Alternatively install {.pkg doconv} to convert with Microsoft Word."
    ),
    call = call
  )
}
