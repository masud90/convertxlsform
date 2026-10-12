# convertxlsform rewrite plan

Status: approved 2026-10-12. Branch: `rewrite`. Target release: 0.3.0.

## Goal

Anyone can install from GitHub and turn an XLSForm into a Word or PDF paper version with one line, from RStudio, any IDE, or a terminal:

```r
remotes::install_github("masud90/convertxlsform")
convertxlsform::convertxlsform("form.xlsx")                  # form.docx next to the input
convertxlsform::convertxlsform("form.xlsx", format = "pdf")  # form.pdf
```

```bash
Rscript -e 'convertxlsform::cli()' form.xlsx --format pdf --lang bn
```

The package must pass `R CMD check --as-cran` with no errors or warnings, since a CRAN submission is planned.

## Settled decisions (from CLAUDE.md, 2026-10-12)

- Keep `%>%` (magrittr).
- New argument names (`lang`, `nested_serial_numbers`, ...), no v0.2.1 compatibility layer.
- CRAN submission planned, so the design keeps hard dependencies on CRAN, avoids system requirements in Imports, and keeps examples fast and offline.

## Scope: start from scratch

All current `R/*.R`, `tests/`, `man/` and `NAMESPACE` are replaced. Kept: DESCRIPTION metadata, LICENSE, `inst/extdata/multilang/`, CI workflow, README/NEWS (rewritten), CLAUDE.md (updated). The v0.2.1 monolith (`git show main:R/convertxlsform.R`) is a reference for layout choices only, not code to copy.

## Public API

| Function | Purpose |
|---|---|
| `convertxlsform(xlsform, output = NULL, format = NULL, lang = NULL, ...)` | One-line conversion. Returns the output path invisibly. `format` is inferred from `output`'s extension, else `"docx"`. |
| `xlsform_languages(xlsform)` | Lists languages found in the form, with the default. |
| `xlsform_outline(xlsform, lang = NULL, ...)` | Returns the parsed, numbered form as a tibble-like data frame (one row per displayed item). Useful for inspection and for tests. |
| `docx_to_pdf(docx, pdf = NULL)` | Converts a .docx to PDF via LibreOffice (or Word, see PDF route). |
| `cli()` | Terminal entry point: parses `commandArgs()` by hand (no optparse). |

Main options for `convertxlsform()` (all with defaults):
`lang` (NULL = settings `default_language`, else first language found), `nested_serial_numbers = TRUE`, `include_relevants = TRUE`, `include_constraints = TRUE`, `include_hints = TRUE`, `include_choice_name = TRUE` (show codes), `show_variable_names = FALSE`, `include_hidden = FALSE` (calculate, metadata), `media_root = dirname(xlsform)`, `title = NULL` (override form_title), `page_size = c("A4", "letter")`, `orientation = c("portrait", "landscape")`, `margins_in = 1`, `font_family = "Arial"`, `font_size = 11`, `template = NULL` (user reference .docx), `max_choices = Inf` (long lists summarized), `overwrite = FALSE`, `verbose = TRUE`.

## Architecture

One internal document model sits between parsing and rendering, so the Word renderer is simple and a future renderer (HTML, native PDF) can reuse it.

```
read_xlsform()  ->  validate  ->  resolve language  ->  parse types/params
     -> build_structure() (groups, repeats, depth, serials)
     -> build_model()     (list of blocks: title, section, question + answer spec)
     -> render_docx()     (officer + flextable)
     -> docx_to_pdf()     (optional)
```

| File | Contents |
|---|---|
| `R/convertxlsform.R` | main function, argument checks, output path, timing message |
| `R/read.R` | read survey/choices/settings sheets; case-insensitive sheet and column names; drop blank rows; keep Excel row numbers |
| `R/validate.R` | required sheets/columns, unknown types, missing choice lists, unmatched begin/end, duplicate names; errors and warnings cite sheet, column and Excel row |
| `R/language.R` | detect languages from `label::English (en)`, `label::en`, `label::English`; match `lang` given as code, name or full header; fall back to default_language, then plain `label`, then first language, with a message |
| `R/types.R` | type registry: every XLSForm type mapped to a display kind (select, text, number, date, geo, media, note, hidden, ...); `or_other`; type arguments; `parameters` parsing |
| `R/structure.R` | vectorized depth and serial numbering; groups take a number, children nest (1, 1.1, 1.2, 2); `field-list` groups; repeats flagged with `repeat_count` |
| `R/expressions.R` | humanize `relevant`, `constraint`, `choice_filter`, `repeat_count`: `${q}` becomes the question number and name, `selected(${q}, 'a')` becomes `q includes "Label a"`, `.` becomes "this answer", `and`/`or`/`not()` kept readable; unparseable input shown verbatim |
| `R/model.R` | build the block list: one record per displayed item with label, hint, required flag, serial, answer spec (choices, box lines, ranges, units), relevance text, media paths |
| `R/choices.R` | choice lists from the choices sheet, external CSV/XML for `*_from_file` (if found under `media_root`), range choices from start/end/step, `max_choices` summarizing |
| `R/media.R` | resolve `image`/`media::image`/`image::<lang>` files under `media_root`; read PNG/JPEG pixel size from file headers in base R (no magick); scale to max height (question 2 in, choice 0.6 in); warn once per missing file |
| `R/render_docx.R` | officer document from template, page setup, title and front matter (form title, version, language, legend), headings for groups, question paragraphs, answer layouts |
| `R/answer_layouts.R` | one function per display kind: choice lists with hollow circle/square symbols, bordered write-in boxes, numeric boxes with range/constraint, date `__/__/____`, time `__:__`, geopoint lat/long/accuracy, media placeholders, rank boxes, acknowledge tick box |
| `R/pdf.R` | find LibreOffice (`soffice`) on macOS, Windows, Linux (PATH and standard install paths) and convert headless into a temp dir; optional Word route via `doconv` (Suggests); clear error with install instructions if neither found |
| `R/cli.R` | argument parsing, `--help`, exit codes |
| `R/utils.R`, `R/messages.R` | `%||%`, ASCII-safe symbols as `\uXXXX`, `cli`-styled messages |
| `inst/templates/convertxlsform.docx` | reference doc defining styles (Title, Heading 1-3, Question, Hint, Logic, Choice) so users can restyle by passing `template` |

Imports (all CRAN, no system requirements): `readxl`, `officer`, `flextable`, `rlang`, `cli`, `magrittr`, `utils`, `tools`. Dropped: dplyr, tidyr, purrr, stringr, tibble, optparse, magick. Suggests: `testthat (>= 3.0.0)`, `writexl`, `withr`, `doconv`, `knitr`, `rmarkdown`.

## PDF route (approved 2026-10-12)

The PDF is the Word document converted by LibreOffice (free, all platforms), with Microsoft Word as a fallback through `doconv` where installed. One renderer keeps Word and PDF identical. PDF output then requires LibreOffice or Word on the user's machine; Word output requires nothing beyond R packages. Rejected: a second native PDF renderer (LaTeX or HTML + Chrome) doubles the layout work and handles Bengali and other non-Latin scripts less reliably.

## Paper form layout

- Front matter: form title, version and form_id, language, generation date, legend (`*` required, circle = choose one, square = choose all that apply, how skip logic is shown).
- Groups: numbered headings by depth; repeats get a "Repeat for each ..." banner with the humanized `repeat_count`.
- Question: `serial. Label *`, hint in italics, optional `[variable_name]`, then the answer layout, then "Ask if: ..." and "Constraint: ..." lines in a muted style.
- Notes unnumbered, hidden types (calculate, start, end, deviceid, ...) skipped unless `include_hidden = TRUE`.
- Basic markdown in labels (`**bold**`, `*italic*`) rendered as runs; HTML tags stripped.
- Choice codes shown as `(1)` after the label when `include_choice_name = TRUE`.

## Testing

- Unit tests per file (type registry, language matching, parameter parsing, expression humanizing, numbering, image size reader).
- Snapshot tests (`expect_snapshot`) of `xlsform_outline()` for the example form in both languages.
- docx content tests via `officer::read_docx()` + `docx_summary()`: title present, question count, choice text, images embedded.
- PDF test runs only when LibreOffice is found (`skip_if_not(has_soffice())`, `skip_on_cran()`).
- Fixtures built in code (`helper-fixtures.R`) plus two shipped examples: `multilang/` (existing) and a new `all_types/` form covering every XLSForm type, generated by `dev/make_examples.R`.
- CI: R-CMD-check on macOS, Windows, Ubuntu (release, devel, oldrel-1); LibreOffice installed on the Ubuntu job so the PDF path is exercised.

## Documentation

README (install, one-liners for RStudio and terminal, screenshots of a rendered page), vignette "Getting started", function reference with runnable examples using `system.file()` and `tempdir()`, NEWS, pkgdown site deployed by GitHub Actions to GitHub Pages.

## Milestones and delegation

Opus plans, writes the model contract (this file plus `dev/model-spec.md`) and reviews each milestone. Workers are Sonnet (medium) for code and Haiku (medium) for mechanical and doc tasks. Fable is consulted at two checkpoints only.

| # | Milestone | Worker | Depends on | Exit check |
|---|---|---|---|---|
| M0 | Clean slate: remove old R/, tests, man; new DESCRIPTION; utils, messages, template docx stub | Haiku | approval | `R CMD build` runs |
| M1 | read, validate, language, types, structure + tests | Sonnet | M0 | unit tests pass |
| M1b | `dev/make_examples.R` and `inst/extdata/all_types/` | Haiku | M0 (parallel with M1) | form loads in readxl |
| M2 | expressions, choices, model, `xlsform_outline()` + snapshot tests | Sonnet | M1 | snapshots reviewed by Opus |
| Fable 1 | Review model spec and layout against the paper-form use case | Fable | M2 | advice folded in |
| M3 | render_docx, answer layouts, media, template + docx tests | Sonnet | M2 | Opus inspects rendered pages |
| M4 | pdf.R, `convertxlsform()`, `cli()` + tests (parallel with M3 for pdf.R) | Sonnet | M2 | end-to-end on both examples |
| M5 | README, vignette, NEWS, pkgdown, CI matrix, CLAUDE.md update | Haiku | M4 | pkgdown builds |
| M6 | `R CMD check --as-cran` clean, fix findings | Sonnet | M5 | 0 errors, 0 warnings, notes explained |
| Fable 2 | Final review of code and outputs before PR | Fable | M6 | findings fixed |
| M7 | Commit, push `rewrite`, PR to `main`, tag v0.3.0 after merge | Opus (with approval) | M6 | install_github works from a clean library |

## Owner decisions (2026-10-12)

1. PDF: convert the Word output with LibreOffice, Word via `doconv` as fallback.
2. LibreOffice installed locally for PDF checks.
3. Commit locally after each reviewed milestone; push and PR to `main` only with explicit approval.
