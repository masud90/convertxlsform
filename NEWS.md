# convertxlsform 0.3.0

A complete rewrite. The package writes Word and PDF paper versions of an XLSForm through a new API, a terminal interface and one document model that both outputs share.

## New features

- `convertxlsform()` writes a Word document (`.docx`) by default. `format = "pdf"` or `format = c("docx", "pdf")` also writes a PDF. The output location, format and overwrite rules follow from the arguments, and the function returns the paths invisibly.
- PDF output converts the Word document with LibreOffice. Microsoft Word is used through the `doconv` package when LibreOffice is not found. `docx_to_pdf()` and `has_libreoffice()` are exported.
- `cli()` runs a conversion from a terminal, for example `Rscript -e 'convertxlsform::cli()' form.xlsx --format pdf`. Scripts can use `quit(status = convertxlsform::cli())` to get the exit status.
- `xlsform_outline()` returns the numbered form as a data frame, for checking a form before conversion.
- `xlsform_languages()` lists the languages in a form and marks the default.
- Language columns are read from `label::English (en)`, `label::en` and `label::English`. The `lang` argument accepts a code, a name or the full suffix. A blank translated cell falls back to the plain column, then to the default language. Headers with a common language name only, such as `label::French`, also match their ISO 639-1 code (`lang = "fr"`).
- Skip logic, constraints and choice filters are written in plain language. Expressions that cannot be read are shown as written.
- Connective phrases such as "Ask if" come from built-in English and French tables. The `strings` option replaces any of them for another language.
- Paper layouts for each question type: single and multiple choice, rank, long lists as code boxes, write-in boxes, numeric, date and time formats, geographic and media questions, and acknowledgments.
- Groups and repeat groups, with nested numbering (`nested_serial_numbers`) and blank copies of repeats (`repeat_copies`).
- Question images are scaled to at most 2 inches tall, and choice images to at most 0.6 inches. A missing media or choice file triggers one combined warning, and a missing image prints a placeholder.
- Complex-script fonts are chosen from the text for Bengali, Arabic, Devanagari and Ethiopic. The `font_family_cs` option overrides the choice.
- Page size, orientation, margins (up to 2 inches), fonts, a reference `template`, `title`, and switches for hints, constraints, relevants, choice codes, variable names and hidden rows.
- Bundled examples: `inst/extdata/multilang` (multilingual, with an external choice list and images) and `inst/extdata/all_types` (every question type and feature the package renders).

## Changes from 0.2.1

- Arguments are renamed: `xlsform_path` is now `xlsform`, `output_docx` is now `output`, `selected_language` is now `lang`, and `number_questions` is now `nested_serial_numbers`. The 0.2.1 names are not accepted.
- HTML tags in labels are stripped. Emphasis in Markdown is rendered.
- Imports are reduced to `cli`, `flextable`, `magrittr`, `officer`, `readxl`, `rlang` and `zip`, plus base packages (`stats`, `tools`, `utils`). `optparse` is no longer used, and the command line is parsed by `cli()`.

## Known issues

- PDF output can show blank text for Bengali and other complex scripts when LibreOffice cannot find a suitable font on the machine. The Word document is correct. A fix is planned for 0.4.0.

# convertxlsform 0.2.1 (2025-07-XX)

- Title field capitalization corrected in DESCRIPTION.
- Added runnable examples in documentation.
- Version bump for CRAN resubmission.

# convertxlsform 0.2.0 (2025-07-19)

**First public release.**

## Highlights

-   **Full XLSForm support**
    -   Reads `survey`, `choices`, and `settings` sheets; errors if any are missing or malformed.\
    -   Preserves HTML/Markdown in `label`, `hint`, and `constraint_message` via `commonmark` to `officer::body_add_html`.
-   **Rich Word output**
    -   Document title (from `form_title`) in 16 pt bold Arial, centered.\
    -   “Mandatory questions are marked with an asterisk (\*) symbol.” note under the header.\
    -   Optional language line.\
    -   All body text in 11 pt Arial, left-aligned.
-   **Question formatting**
    -   Required questions flagged with `*`.\
    -   Optional nested numbering (`number_questions = TRUE`) for top‑level, groups, and repeats.\
    -   Notes printed as `Note: …` (no numbering).\
    -   Hints and images (capped at 2″ height) embedded inline.\
    -   Constraint messages in parentheses; relevant clauses appended.
-   **Answer layouts**
    -   **select_one / rank**: larger hollow‐circle bullets (`\u25EF`).\
    -   **select_multiple**: larger hollow‐square bullets (`\u2B1B`).\
    -   **select_one_from_file**: prints attachment notice.\
    -   **range**: generates choices from `start`, `end`, `step`.\
    -   **Text / numeric / date / time**: bordered text boxes spanning full page width (6.5″), with 2–4 empty lines.\
    -   **Media / file / note types**: appropriate prompts or skips.
-   **Choice images**
    -   If `choices$image` is provided and file exists, embeds under each bullet (max 1″ height).\
    -   Missing files trigger a warning and are noted next to the label.

## Packaging & internals

-   **Dependencies**: `readxl`, `officer`, `commonmark`, `flextable`, `magick` declared in Imports.\
-   **ASCII compliance**: all Unicode in code replaced with `\uXXXX` escapes.\
-   **Roxygen imports**: explicit `@importFrom` for all used functions.\
-   **Tests**: initial `testthat` scaffold with basic validation tests.

------------------------------------------------------------------------

Thank you for using **convertxlsform**: feedback and issues are welcome on GitHub (<https://github.com/masud90/convertxlsform>)!
