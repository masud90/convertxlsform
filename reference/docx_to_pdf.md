# Convert a Word document to PDF

Converts a `.docx` file to PDF with LibreOffice, or with Microsoft Word
through the `doconv` package. LibreOffice is free and runs on all
platforms. Each call uses a fresh temporary LibreOffice profile, so it
works while a LibreOffice window is open and in parallel runs.

## Usage

``` r
docx_to_pdf(
  docx,
  pdf = NULL,
  engine = c("auto", "libreoffice", "word"),
  overwrite = FALSE,
  verbose = TRUE,
  fonts = NULL
)
```

## Arguments

- docx:

  Path to an existing `.docx` file.

- pdf:

  Path of the PDF to write. `NULL` (default) uses the path of `docx`
  with the extension changed to `.pdf`.

- engine:

  `"auto"` (default) tries LibreOffice, then Word through `doconv` if
  that package is installed. `"libreoffice"` and `"word"` force one
  engine.

- overwrite:

  Overwrite an existing `pdf`? Default `FALSE`.

- verbose:

  Print progress messages? Default `TRUE`.

- fonts:

  Font files (`.ttf`, `.otf`, `.ttc`) or folders of font files that
  LibreOffice should use for this conversion, in addition to the fonts
  it finds itself. Use it when a script shows as blank text or boxes in
  the PDF although the font is on the machine. `NULL` (default) uses the
  option `convertxlsform.fonts`, if set. Ignored by the Word engine.

## Value

The path of the PDF, invisibly.

## Details

If LibreOffice is installed in an unusual place, set the option
`convertxlsform.soffice` or the environment variable
`CONVERTXLSFORM_SOFFICE` to the path of the `soffice` binary.

## Examples

``` r
if (interactive() && has_libreoffice()) {
  docx <- tempfile(fileext = ".docx")
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "Hello")
  print(doc, target = docx)
  pdf <- docx_to_pdf(docx, pdf = tempfile(fileext = ".pdf"))
  file.exists(pdf)
}
```
