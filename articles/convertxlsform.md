# Getting started with convertxlsform

convertxlsform reads an XLSForm, the questionnaire format used by Open
Data Kit (ODK) and KoboToolbox, and writes a paper version as a Word
document or a PDF. This vignette walks through the main steps using the
example form that ships with the package. All output goes to
[`tempdir()`](https://rdrr.io/r/base/tempfile.html), so nothing is
written next to the installed package.

## The example form

The example is `all_types.xlsx`. It covers the question types and
features the package renders, including groups, repeats, cascading
choice lists, and geo and media questions.

``` r

library(convertxlsform)

path <- system.file("extdata", "all_types", "all_types.xlsx",
                    package = "convertxlsform")
basename(path)
#> [1] "all_types.xlsx"
```

## Languages

[`xlsform_languages()`](https://masud90.github.io/convertxlsform/reference/xlsform_languages.md)
lists the languages in the form and marks the default. Language columns
are named like `label::English (en)` or `label::French (fr)`.

``` r

xlsform_languages(path)
#>         suffix code    name default
#> 1 English (en)   en English    TRUE
#> 2  French (fr)   fr  French   FALSE
```

The default language is used when no `lang` is given. Pass a code
(`"fr"`), a name (`"French"`) or the full suffix (`"French (fr)"`) to
choose another one.

## Inspecting the form

[`xlsform_outline()`](https://masud90.github.io/convertxlsform/reference/xlsform_outline.md)
returns one row per item on the paper form, with the question numbers,
the question kind and the skip logic in plain language. Use it to check
a form before converting it. The example below shows the first ten rows.

``` r

outline <- xlsform_outline(path)
nrow(outline)
#> [1] 43
head(outline[, c("serial", "block", "kind", "name", "relevant_text")], 10)
#>    serial    block        kind       name relevant_text
#> 1    <NA>     note        note intro_note          <NA>
#> 2       1 question  select_one    consent          <NA>
#> 3       2 question        text  follow_up          <NA>
#> 4       3  section begin_group         hh          <NA>
#> 5     3.1 question        text    hh_code          <NA>
#> 6     3.2 question      number    hh_size          <NA>
#> 7     3.3 question      number    land_ha          <NA>
#> 8     3.4 question        date visit_date          <NA>
#> 9     3.5 question        time visit_time          <NA>
#> 10    3.6 question    datetime  last_rain          <NA>
```

Groups and notes also appear as rows. Notes are not numbered, so their
`serial` is `NA`.

## Converting to Word

[`convertxlsform()`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)
writes a Word document. The default output is a `.docx` file with the
same base name as the form. Here the output goes to
[`tempdir()`](https://rdrr.io/r/base/tempfile.html), and the function
returns the path of the file it wrote.

``` r

docx <- convertxlsform(path, output = tempdir(), overwrite = TRUE,
                       verbose = FALSE)
basename(docx)
#> [1] "all_types.docx"
file.exists(docx)
#> [1] TRUE
```

To produce the French version, pass `lang`. Where a French hint is
blank, the English hint is shown, because a blank cell falls back to the
plain column and then to the default language.

``` r

docx_fr <- convertxlsform(path, output = file.path(tempdir(), "all_types_fr.docx"),
                          lang = "fr", overwrite = TRUE, verbose = FALSE)
basename(docx_fr)
#> [1] "all_types_fr.docx"
```

## Converting to PDF

A PDF is made by converting the Word document with LibreOffice, or with
Microsoft Word through the `doconv` package. PDF output needs
LibreOffice (or Word with `doconv`), so the code below is shown but not
run when the vignette is built.

``` r

pdf <- convertxlsform(path, output = tempdir(), format = "pdf",
                      overwrite = TRUE, verbose = FALSE)
basename(pdf)
file.exists(pdf)
```

If LibreOffice is installed in an unusual place, set
`options(convertxlsform.soffice = "/path/to/soffice")` before
converting.

## Next steps

- The README lists every option, the terminal commands and how to
  translate the built-in phrases with `strings`.
- [`?convertxlsform`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)
  gives the full description of the options.
- [`?xlsform_outline`](https://masud90.github.io/convertxlsform/reference/xlsform_outline.md)
  describes the columns of the outline.
