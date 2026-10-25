# Changing the look with a template

convertxlsform writes the paper form into a reference Word document, the
template, and uses the paragraph styles it defines. Pass your own
template with the `template` option (`--template` in the terminal) to
change the spacing, indents, shading and borders of the form to match
your organization’s house style. This vignette shows where the default
template lives, which styles it defines, how to change them in Word,
LibreOffice or R, and which options always win over the template. All
files are written to
[`tempdir()`](https://rdrr.io/r/base/tempfile.html).

## The default template

The package ships one template. It is A4 with one-inch margins and Arial
11 point text.

``` r

library(convertxlsform)

default_template <- system.file("templates", "convertxlsform.docx",
                                package = "convertxlsform")
basename(default_template)
#> [1] "convertxlsform.docx"
```

The template is built by the script
[`dev/make_template.R`](https://github.com/masud90/convertxlsform/blob/main/dev/make_template.R)
in the source repository. Read it to see every value the default styles
use.

Start your own template from a copy of the default one, so that every
style the package looks for is already there.

``` r

my_template <- file.path(tempdir(), "my_template.docx")
file.copy(default_template, my_template, overwrite = TRUE)
#> [1] TRUE
```

## The styles convertxlsform uses

Each part of the paper form is written with one paragraph style. The
style names are the names shown in Word and LibreOffice.

| Style | Used for |
|----|----|
| `Title` | The form title at the top of the first page. |
| `heading 1`, `heading 2`, `heading 3` | Group headings, by nesting level. Deeper groups use `heading 3`. |
| `cxf Question` | The question number and label. |
| `cxf Hint` | Hints, the form details under the title, and the legend. |
| `cxf Logic` | Skip logic (“Ask if”), constraints and “Go to” lines. |
| `cxf Note` | Notes (questions of type `note`). |
| `cxf Banner` | The “For each” banner at the start of a repeat. |
| `Normal` | The header and footer, and any style that is missing. |

The answer options, answer lines and grids are Word tables, so they do
not take a paragraph style. The template also defines `cxf Choice`,
which the current version does not use.

If your template lacks one of these styles, that part of the form uses
`Normal` with the package’s built-in spacing, so a template made from
scratch still works. Keep the style names exactly as written above,
including spaces and capital letters.

## What a style controls

convertxlsform sets the font and size of all text directly, from the
`font_family` and `font_size` options, and sets the colour of headings,
hints, skip logic and the repeat banner. Those settings win over the
template. What the template controls is the paragraph:

- space before and after, and line spacing;
- indents and alignment;
- shading and borders;
- “keep with next” and other page-break rules;
- the text colour of questions and notes.

So to change the font of the whole form, use `font_family`. To put a
light rule above each question and give the repeat banner a different
background, change the `cxf Question` and `cxf Banner` styles.

## Changing the styles in Word or LibreOffice

In Microsoft Word:

1.  Open your copy of the template.
2.  Open the Styles pane (Home tab, the small arrow in the corner of the
    Styles group).
3.  Right-click a style, for example `cxf Question`, and choose
    **Modify**.
4.  Use **Format \> Paragraph** for spacing and indents, and **Format \>
    Border** for borders and shading.
5.  Save the file as a Word document (`.docx`).

In LibreOffice Writer:

1.  Open your copy of the template.
2.  Press F11 to open the Styles deck and choose Paragraph Styles (with
    “All Styles” in the list at the bottom).
3.  Right-click a style and choose **Edit Style**. The Indents &
    Spacing, Borders and Area tabs hold the settings.
4.  Save with **File \> Save As**, keeping the Word 2007-365 (`.docx`)
    format.

Do not rename the styles, and delete the sample text if you add any: the
body of the template is not copied into the form.

## Changing the styles in R

The officer package, which convertxlsform uses to write Word files, can
change the styles too. The code below gives each question a thin grey
rule above it and more space, makes question labels dark blue, and gives
the repeat banner a pale yellow background.

``` r

library(officer)

tpl <- read_docx(my_template)
tpl <- docx_set_paragraph_style(
  tpl, style_id = "cxfQuestion", style_name = "cxf Question",
  base_on = "Normal",
  fp_p = fp_par(keep_with_next = TRUE, padding.top = 12, padding.bottom = 4,
                border.top = fp_border(color = "#BFBFBF", width = 0.5)),
  fp_t = fp_text_lite(color = "#00205B")
)
tpl <- docx_set_paragraph_style(
  tpl, style_id = "cxfBanner", style_name = "cxf Banner",
  base_on = "Normal",
  fp_p = fp_par(shading.color = "#FFF2CC", keep_with_next = TRUE,
                padding.top = 6, padding.bottom = 4),
  fp_t = fp_text_lite(bold = TRUE)
)
print(tpl, target = my_template)
```

Check that the styles are there with
[`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.html).

``` r

st <- styles_info(read_docx(my_template))
st[startsWith(st$style_name, "cxf"), c("style_id", "style_name")]
#>       style_id   style_name
#> 26 cxfQuestion cxf Question
#> 27     cxfHint     cxf Hint
#> 28    cxfLogic    cxf Logic
#> 29   cxfChoice   cxf Choice
#> 30     cxfNote     cxf Note
#> 31   cxfBanner   cxf Banner
```

## Converting with your template

Pass the template with `template`. The example form ships with the
package.

``` r

path <- system.file("extdata", "all_types", "all_types.xlsx",
                    package = "convertxlsform")
docx <- convertxlsform(path,
                       output = file.path(tempdir(), "all_types_custom.docx"),
                       template = my_template, overwrite = TRUE,
                       verbose = FALSE)
basename(docx)
#> [1] "all_types_custom.docx"
file.exists(docx)
#> [1] TRUE
```

From the terminal, the same conversion is:

``` sh
Rscript -e 'convertxlsform::cli()' all_types.xlsx --template my_template.docx
```

## Options that override the template

These options are applied to every document, whatever the template says:

| Option | Terminal | What it sets |
|----|----|----|
| `font_family` | `--font` | The font of all text (default Arial). |
| `font_size` | `--font-size` | The base size of all text, in points (default 11). |
| `font_family_cs` | `--font-cs` | The font for complex scripts such as Arabic or Bengali. |
| `page_size` | `--page-size` | A4 or letter (default A4). |
| `orientation` |  | Portrait or landscape (default portrait). |
| `margins_in` | `--margins` | All four page margins, in inches (default 1). |

For example, a letter-size form in Calibri 10 point with narrow margins,
keeping the paragraph styles of your template:

``` r

docx_letter <- convertxlsform(
  path, output = file.path(tempdir(), "all_types_letter.docx"),
  template = my_template, page_size = "letter", margins_in = 0.75,
  font_family = "Calibri", font_size = 10, overwrite = TRUE, verbose = FALSE
)
basename(docx_letter)
#> [1] "all_types_letter.docx"
```

## Header, footer and logo

The header (the form title and version) and the footer (the form ID, the
date and the page number) are written by convertxlsform, and they
replace any header or footer in the template. A logo placed in the
template’s header is therefore not kept.

To add a logo, convert the form first, then open the Word file,
double-click the header area and insert the picture (**Insert \>
Pictures** in Word, **Insert \> Image** in LibreOffice). The header text
convertxlsform wrote stays in place, so you can put the logo beside it.
Make the PDF from the edited file, either from Word or LibreOffice with
Save as PDF, or with
[`docx_to_pdf()`](https://masud90.github.io/convertxlsform/reference/docx_to_pdf.md):

``` r

docx_to_pdf(docx)
```

## Next steps

- The “Getting started” vignette covers languages, the outline and PDF
  output.
- [`?convertxlsform`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)
  describes every option.
