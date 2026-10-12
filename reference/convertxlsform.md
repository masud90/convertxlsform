# Convert an XLSForm to a Word or PDF document

Reads an XLSForm (as used by Open Data Kit and KoboToolbox) and writes a
formatted paper version for interviewing and review. The default output
is a Word document next to the input file. A PDF is made by converting
the Word document with LibreOffice (or Microsoft Word through the
`doconv` package); see
[`docx_to_pdf()`](https://masud90.github.io/convertxlsform/reference/docx_to_pdf.md)
and
[`has_libreoffice()`](https://masud90.github.io/convertxlsform/reference/has_libreoffice.md).

## Usage

``` r
convertxlsform(xlsform, output = NULL, format = NULL, lang = NULL, ...)
```

## Arguments

- xlsform:

  Path to an XLSForm (`.xlsx`, `.xlsm` or `.xls`).

- output:

  Output file or folder, see Details. `NULL` (default) writes next to
  `xlsform`.

- format:

  `"docx"`, `"pdf"`, `"both"` or `c("docx", "pdf")`. `NULL` (default)
  takes the format from the extension of `output`, and is `"docx"` when
  there is no such extension.

- lang:

  Language to use: a code (`"fr"`), a name (`"French"`) or the full
  header suffix (`"French (fr)"`). `NULL` (default) uses the
  `default_language` of the settings sheet, or the first language found.
  See
  [`xlsform_languages()`](https://masud90.github.io/convertxlsform/reference/xlsform_languages.md).

- ...:

  Further options, all named. Defaults are in brackets.

  - `nested_serial_numbers` (`TRUE`): number questions inside groups as
    1.1, 1.2; `FALSE` gives flat numbers 1, 2, 3.

  - `include_relevants` (`TRUE`): show skip logic ("Ask if: ...").

  - `include_constraints` (`TRUE`): show constraints.

  - `include_hints` (`TRUE`): show hints and guidance hints.

  - `include_choice_name` (`TRUE`): show the choice code after each
    choice label.

  - `show_variable_names` (`FALSE`): show the variable name of each
    question.

  - `include_hidden` (`FALSE`): include hidden types such as
    `calculate`, `start` and `end`.

  - `media_root` (folder of `xlsform`): where images, audio, video and
    external choice files are looked up.

  - `title` (`NULL`): form title; overrides `form_title` of the settings
    sheet.

  - `page_size` (`"A4"`): `"A4"` or `"letter"`.

  - `orientation` (`"portrait"`): `"portrait"` or `"landscape"`.

  - `margins_in` (`1`): page margins in inches, above 0 and at most 2.

  - `font_family` (`"Arial"`): font for the document.

  - `font_size` (`11`): base font size in points, 6 to 20.

  - `template` (`NULL`): a reference `.docx` whose styles are used.

  - `max_choices` (`30`): choice lists longer than this get a code box
    layout instead of a printed list.

  - `strings` (`NULL`): named list overriding connective phrases such as
    "Ask if".

  - `repeat_copies` (`1`): number of blank copies printed for each
    repeat group, 1 to 20.

  - `font_family_cs` (`NULL`): font for complex scripts; chosen from the
    language when `NULL`.

  - `overwrite` (`FALSE`): replace existing output files.

  - `verbose` (`TRUE`): print progress messages.

## Value

The path or paths of the files written, as a character vector,
invisibly.

## Where the file goes

- `output = NULL`: same folder and base name as `xlsform`, with the new
  extension.

- `output` is an existing folder: `<base name>.<ext>` inside it.

- `output` is a file path: used as given. Its extension (`.docx` or
  `.pdf`) sets `format` when `format` is `NULL`, and must agree with
  `format` otherwise.

- Two formats (`format = c("docx", "pdf")`) write two files with the
  same base name. With a file path in `output`, its extension is
  replaced.

An existing file is never replaced unless `overwrite = TRUE`.

## See also

[`xlsform_outline()`](https://masud90.github.io/convertxlsform/reference/xlsform_outline.md)
to inspect a form before converting it,
[`docx_to_pdf()`](https://masud90.github.io/convertxlsform/reference/docx_to_pdf.md),
and [`cli()`](https://masud90.github.io/convertxlsform/reference/cli.md)
for the terminal.

## Examples

``` r
path <- system.file("extdata", "multilang", "multilang.xlsx",
                    package = "convertxlsform")
out <- file.path(tempdir(), "multilang.docx")
convertxlsform(path, output = out, verbose = FALSE, overwrite = TRUE)

# \donttest{
# Another language, flat question numbers, US letter paper
convertxlsform(path, output = tempdir(), lang = "bn",
               nested_serial_numbers = FALSE, page_size = "letter",
               verbose = FALSE, overwrite = TRUE)

# PDF needs LibreOffice
if (interactive() && has_libreoffice()) {
  convertxlsform(path, output = tempdir(), format = "pdf",
                 verbose = FALSE, overwrite = TRUE)
}
# }
```
