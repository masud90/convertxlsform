# convertxlsform

[![CRAN status](https://www.r-pkg.org/badges/version/convertxlsform)](https://CRAN.R-project.org/package=convertxlsform)

Convert ODK‐style XLSForm questionnaires into neatly formatted Microsoft Word (`.docx`) documents.

---

## Installation

From CRAN:
```r
install.packages("convertxlsform")
```

Or the latest development version on GitHub:

```r
# install.packages("devtools")
devtools::install_github("masud90/convertxlsform")
```
## Usage

```r
library(convertxlsform)
```

### 1. Basic conversion (numbering on by default)

```r
convertxlsform("survey.xlsx")
#> Document generated in 0.42 secs: /full/path/survey.docx
```

### 2. Specify language and output path

```r
convertxlsform(
  xlsform_path      = "survey.xlsx",
  selected_language = "en",
  output_docx       = "reports/MyForm.docx"
)
```

### 3. Turn numbering off

```r
convertxlsform("survey.xlsx", number_questions = FALSE)
```

## Function signature

```r
convertxlsform(
  xlsform_path,        # (chr) Path to XLSForm .xlsx
  selected_language = "en",  # (chr) language code (must match label::<lang> column)
  output_docx       = NULL,  # (chr) destination .docx (defaults to basename(xlsx))
  number_questions  = TRUE   # (bool) nested numbering of questions/groups
)
```

## XLSForm Requirements

- `survey` sheet with `type`, `name`, `label` columns (and optional `label::<lang>`, `hint::<lang>`, `instructions::<lang>`, `required`, `relevant`, `constraint_message`, `parameters`, `repeat_count`).
- `choices` sheet with `list_name`, `name`, `label` column (and optional `label::<lang>`, `image`).
- `settings` sheet with `form_title` column (and optional `form_title::<lang>`).
- If you have image files in the CAPI and wish for them to appear in the PAPI, upload the files of the same name to the root folder of your project (where the CAPI form is).


## Output formatting

- Title: Arial 16 pt, bold, centered.
- Mandatory note: “Mandatory questions are marked with an asterisk (*) symbol.”
- Body text: Arial 11 pt, left-aligned.
- Question numbering: Nested (“1.”, “1.1.” in groups, etc.), default TRUE.
- Notes: Unnumbered, prefixed Note:.
- Select questions:
  - select_one, rank, range: hollow‐circle bullets (◯).
  - select_multiple: hollow‐square bullets (▢).
  - Text/range boxes: bordered tables spanning full width with only blank lines.
- Images:
  - Question images: up to 2″ tall.
  - Choice images: up to 1″ tall, with warnings if missing.

 ## Example

```r
 # Multi‐language support:
convertxlsform("survey_multilang.xlsx", selected_language = "fr")

# Custom output location and no numbering:
convertxlsform(
  xlsform_path     = "survey.xlsx",
  output_docx      = "out/MySurvey.docx",
  number_questions = FALSE
)
```

## Dependencies

- `readxl` (read .xlsx)
- `officer` (generate Word docs)
- `commonmark` (Markdown parsing – unused by default)
- `flextable` (create boxed answer fields)
- `magick` (image sizing)
- `stats` (utility setNames)


## License

This project is licensed under the MIT license.
See [LICENSE](license.md) for full text.

## Contribution

Clone the repo and send a pull request.
