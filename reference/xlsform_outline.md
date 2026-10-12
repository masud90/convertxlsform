# Outline of an XLSForm

Parses an XLSForm and returns one row for each item that would appear on
the paper form: sections, questions, notes and the end of repeat groups.
It shows the question numbers, the answer type, the number of choices,
and the skip logic and constraints in plain language, so a form can be
checked before it is converted.

## Usage

``` r
xlsform_outline(xlsform, lang = NULL, ...)
```

## Arguments

- xlsform:

  Path to an XLSForm (`.xlsx` or `.xls`).

- lang:

  Language to use: a code (`"fr"`), a name (`"French"`) or the full
  header suffix (`"French (fr)"`). `NULL` uses the `default_language` of
  the settings sheet, or the first language found. See
  [`xlsform_languages()`](https://masud90.github.io/convertxlsform/reference/xlsform_languages.md).

- ...:

  Further options: `nested_serial_numbers`, `include_relevants`,
  `include_constraints`, `include_hints`, `include_hidden`,
  `media_root`, `max_choices` (lists longer than this get a code box
  layout), `strings`, `show_variable_names`, `title` and the other
  conversion options. Page and font options are accepted and have no
  effect here.

## Value

A data.frame with one row per block and columns `row` (Excel row of the
form), `serial` (question number, NA when unnumbered), `depth` (number
of enclosing groups), `block` (`"section"`, `"section_end"`, `"note"` or
`"question"`), `kind` (the question kind, for example `"select_one"`, or
`"begin_group"`, `"begin_repeat"`, `"end_repeat"`, `"note"`), `name`,
`label`, `required`, `n_choices` (all choices of a select or rank
question, not counting an "other" row), `relevant_text`,
`constraint_text` (the constraint message when the form has one, else
the humanized expression) and `layout` (for select and rank questions:
`"list"`, `"columns"` or `"code_box"`, from the number of choices and
`max_choices`; NA otherwise). A label may contain line breaks (`\n`).

## Examples

``` r
path <- system.file("extdata", "multilang", "multilang.xlsx",
                    package = "convertxlsform")
out <- xlsform_outline(path)
head(out[, c("serial", "block", "kind", "name", "relevant_text")], 12)
#>    serial    block         kind             name relevant_text
#> 1       1  section  begin_group           groupA          <NA>
#> 2     1.1 question         date       surveydate          <NA>
#> 3     1.2 question   select_one         enumname          <NA>
#> 4    <NA>     note         note     note_consent          <NA>
#> 5     1.3 question  acknowledge          consent          <NA>
#> 6       2  section  begin_group           groupB          <NA>
#> 7     2.1 question       number           hhsize          <NA>
#> 8     2.2  section begin_repeat         hhroster          <NA>
#> 9   2.2.1 question         text         mem_name          <NA>
#> 10  2.2.2 question   select_one          mem_sex          <NA>
#> 11  2.2.3 question         date    mem_birthyear          <NA>
#> 12  2.2.4 question   select_one mem_relationship          <NA>

# Another language, with flat question numbers
out <- xlsform_outline(path, lang = "bn", nested_serial_numbers = FALSE)
head(out[, c("serial", "name")], 5)
#>   serial         name
#> 1   <NA>       groupA
#> 2      1   surveydate
#> 3      2     enumname
#> 4   <NA> note_consent
#> 5      3      consent
```
