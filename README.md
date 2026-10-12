# convertxlsform

Convert ODK-style XLSForm questionnaires into neatly formatted Microsoft Word (`.docx`) documents.

> **Status:** this is the `rewrite` branch (version 0.2.1.9000), a modular rewrite under development. It reads and validates XLSForms but **does not yet write `.docx` files**. To produce Word documents today, use `main` (v0.2.1). Rendering is being ported from v0.2.1 and the rewrite will be merged as 0.3.0.

---

## Installation

The package is not on CRAN. Install from GitHub:

```r
# install.packages("remotes")

# Stable v0.2.1 (writes .docx)
remotes::install_github("masud90/convertxlsform")

# Development branch (rewrite, no .docx output yet)
remotes::install_github("masud90/convertxlsform@rewrite")
```

## Usage

```r
library(convertxlsform)

convertxlsform("survey.xlsx")

convertxlsform(
  xlsform_path = "survey.xlsx",
  output_docx  = "reports/MyForm.docx",
  lang         = "fr"
)
```

A synthetic multilingual example form ships with the package:

```r
system.file("extdata", "multilang", "multilang.xlsx", package = "convertxlsform")
```

## Function signature (rewrite branch)

```r
convertxlsform(
  xlsform_path,                 # (chr) path to the XLSForm .xlsx
  output_docx           = NULL, # (chr) destination .docx
  lang                  = "en", # (chr) language code, matches label::<lang> columns
  include_relevants     = TRUE, # (lgl) show skip logic (relevant)
  include_choice_name   = TRUE, # (lgl) show choice names next to labels
  nested_serial_numbers = TRUE, # (lgl) nested numbering of questions and groups
  page_size             = "A4", # (chr) page size
  margins_in            = 1,    # (num) page margins in inches
  media_root            = ".",  # (chr) folder holding question and choice media
  verbose               = TRUE  # (lgl) print progress messages
)
```

Returns the output path invisibly. Language columns resolve in this order: `label::<lang>`, plain `label`, then `label::en`.

The v0.2.1 signature on `main` is `convertxlsform(xlsform_path, selected_language = "en", output_docx = NULL, number_questions = TRUE)`. Whether the rewrite keeps those argument names is an open decision.

## XLSForm requirements

- `survey` sheet with `type`, `name`, `label` columns (and optional `label::<lang>`, `hint::<lang>`, `instructions::<lang>`, `required`, `relevant`, `constraint_message`, `parameters`, `repeat_count`).
- `choices` sheet with `list_name`, `name`, `label` columns (and optional `label::<lang>`, `image`).
- `settings` sheet with `form_title` column (and optional `form_title::<lang>`).
- If you have image files in the CAPI form and want them to appear in the PAPI version, put files of the same name in the root folder of your project (where the CAPI form is).

## Output formatting (v0.2.1 behavior; being ported)

- Title: Arial 16 pt, bold, centered.
- Mandatory note: "Mandatory questions are marked with an asterisk (*) symbol."
- Body text: Arial 11 pt, left-aligned.
- Question numbering: nested ("1.", "1.1." in groups, etc.), on by default.
- Notes: unnumbered, prefixed "Note:".
- Select questions:
  - select_one, rank, range: hollow-circle bullets.
  - select_multiple: hollow-square bullets.
  - Text and range boxes: bordered tables spanning the full width with only blank lines.
- Images:
  - Question images: up to 2 in tall.
  - Choice images: up to 1 in tall, with warnings if missing.

## Dependencies

Imports (from `DESCRIPTION`): `dplyr`, `magrittr`, `optparse`, `purrr`, `readxl`, `rlang`, `stringr`, `tibble`, `tidyr`. Suggests: `testthat`, `writexl`.

The rendering packages used by v0.2.1 (`officer`, `flextable`, `magick`) are not yet imported on this branch.

## License

This project is licensed under the MIT license. See [LICENSE.md](LICENSE.md) for the full text.

## Contribution

Clone the repo and send a pull request. Develop on the `rewrite` branch.
