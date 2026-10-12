# List the languages of an XLSForm

Finds the languages used in the `label::<language>` style columns of the
survey and choices sheets, for example `label::English (en)`,
`label::en` or `label::English`.

## Usage

``` r
xlsform_languages(xlsform)
```

## Arguments

- xlsform:

  Path to an XLSForm (`.xlsx` or `.xls`).

## Value

A data.frame with one row per language and columns `suffix` (the text
after `::`), `code`, `name` and `default` (TRUE for the language used
when no `lang` is requested: the `default_language` of the settings
sheet if it matches, otherwise the first language). The data.frame has
no rows for a single-language form.

## Examples

``` r
path <- system.file("extdata", "multilang", "multilang.xlsx",
                    package = "convertxlsform")
xlsform_languages(path)
#>         suffix code    name default
#> 1 English (en)   en English    TRUE
#> 2 Bengali (bn)   bn Bengali   FALSE
```
