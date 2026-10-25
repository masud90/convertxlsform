# convertxlsform 0.4.0

## New features

- "Go to" instructions. A single-choice question whose answer hides later questions now shows, next to that choice, the question to go to next, for example "No → Go to Q5", or "Go to end". The instruction is worked out from the `relevant` expressions of the questions that follow; it is left out when the result depends on other answers or would leave a repeat group. `include_skips = FALSE` (or `--no-skips`) turns it off.
- Roster tables for repeat groups. A repeat of simple questions (text, numbers, single and multiple choice, dates, times, barcodes, ranges) is drawn as a table with one row per member and one column per question, with the code lists and skip logic below it. A fixed `repeat_count` sets the number of rows, otherwise `roster_rows` (default 5). `repeat_layout = "auto"` (the default) uses a roster for repeats of up to 6 such questions, `"roster"` allows up to 10, and `"copies"` keeps the earlier layout.
- Grids for `table-list` and `list-nolabel`. A group with the `table-list` appearance whose questions share a choice list, and a `label` question followed by `list-nolabel` questions, are drawn as one grid with a row per question and a column per choice. The `label` question gives the column headings and takes no number.
- Bilingual documents. `lang2` (`--lang2`) prints a second language on a new line under each label, hint, constraint message and choice of the first.
- Right-to-left layout. Forms in Arabic or Hebrew are laid out right to left: paragraphs, tables and the page. `rtl = TRUE` or `FALSE` (`--rtl`, `--ltr`) overrides the detection.
- The `entities` sheet of ODK entity forms is read. The entity lists are named in the front matter, and with `show_variable_names = TRUE` a question's `save_to` property is shown next to its name.
- More XLSForm types and spellings: `background-geopoint` (hidden), `osm` (captured on the device), and the pyxform aliases `select all that apply`, `select_all_that_apply`, `select1`, `int`, `string`, `gps`, `location` and `photo`.
- More XLSForm functions are written in plain language in skip logic, constraints and choice filters: `string-length()`, `regex()`, `selected-at()`, `jr:choice-name()`, `count()`, `position(..)`, `indexed-repeat()`, `if()`, `coalesce()`, `int()`, `round()`, `number()`, `date()` and `decimal-date-time()`. An age check such as `int((today() - ${dob}) div 365.25) >= 18` reads "years since Q3 (dob) >= 18". `once()`, `uuid()` and `pulldata()` are shown as "a calculated value". The wording is in the string tables, so it can be translated or overridden with `strings` (#20).
- XML choice files are read for `select_one_from_file` and `select_multiple_from_file`, as CSV files already were (#21).
- Language names are recognized in many more headers without a code, in English or in their own script, for example `label::Spanish`, `label::Español` or `label::বাংলা`, so the built-in strings of that language are used. `lang = "Arabic"` also finds a column headed `label::ar` (#18).
- New terminal options: `--lang2`, `--no-skips`, `--rtl`, `--ltr`, `--font-cs`, `--repeat-layout` and `--roster-rows`.

## Bug fixes

- PDF output no longer shows blank text for Bengali, Ethiopic and other complex scripts when LibreOffice does not find the font on its own (the known issue of 0.3.0). The files of the complex-script font are copied into the temporary LibreOffice profile used for the conversion, and a warning names the font when it is not installed. `docx_to_pdf()` gains a `fonts` argument, and the option `convertxlsform.fonts`, to give LibreOffice more font files or folders.
- The complex-script font is now the first installed font from a list for each script (for Bengali: Noto Sans Bengali, Noto Serif Bengali, Kohinoor Bangla, Bangla Sangam MN, Nirmala UI, Vrinda, and others), instead of always the Noto font.
- The last paragraph of a question no longer keeps with the next question. The template styles keep paragraphs together, and officer did not write the paragraph-level override, so long runs of questions were glued together and page breaks fell in odd places.

## Performance

- Rendering is several times faster and now grows linearly with the size of the form. Answer boxes and choice lists are written as Word XML directly instead of through flextable, and the document body is written in one pass. In a test form of 300 questions, writing the Word document took about 6 seconds instead of about 2 minutes.

## Other changes

- The DESCRIPTION Title and Description put single quotes only around software names ('Microsoft Word', 'LibreOffice', 'Open Data Kit', 'KoboToolbox'), as CRAN asked.
- The README has a CRAN status badge and gives `install.packages("convertxlsform")` as the main way to install.
- `systemfonts` is a new import (it was already installed with flextable). officer 0.6.10 or later is required.

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

- PDF output can show blank text for Bengali and other complex scripts when LibreOffice cannot find a suitable font on the machine. The Word document is correct. Fixed in 0.4.0.

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
