# convertxlsform

## Purpose and audience

R package that converts XLSForm (ODK / KoboToolbox) questionnaires into Word documents (paper versions for review and PAPI). Author and maintainer: Masud Rahman, regional statistician, UNHCR RBESA, Nairobi. Users are survey managers and information management staff who need a reviewable paper copy of a CAPI instrument.

## Repo and branches

- `main` (origin https://github.com/masud90/convertxlsform): v0.2.1, single-file implementation in `R/convertxlsform.R` (about 600 lines). It already writes .docx.
- `rewrite` (current): modular rewrite, version 0.2.1.9000. Develop here. Do not commit or push unless asked.
- Workflow: open a PR to `main` once the rewrite reaches parity with v0.2.1 (writes a .docx with at least the v0.2.1 features). Bump to 0.3.0 at merge.
- Reference for porting rendering: `git show main:R/convertxlsform.R` (and `git show main:README.md`).
- The v0.2.1 README claimed a CRAN release, but CRAN has no record of the package. Do not state or imply it is on CRAN.

## Architecture

Pipeline by phase. Data flow: `survey` and `choices` tibbles -> structure tibble -> .docx.

| Phase | File | Role |
|---|---|---|
| 1 | `R/convertxlsform.R` | Entry point `convertxlsform(xlsform_path, output_docx = NULL, lang = "en", include_relevants = TRUE, include_choice_name = TRUE, nested_serial_numbers = TRUE, page_size = "A4", margins_in = 1, media_root = ".", verbose = TRUE)`. Returns the output path invisibly. Currently writes a placeholder tempfile, no .docx yet. |
| 2 | `R/read_validate.R` | Read sheets, validate required sheets and columns, normalize names, resolve language columns (`label::<lang>`, then plain `label`, then `label::en`), parse `base_type`, `type_arg`, `parameters`, `required`. |
| 3 | `R/structure.R` | Group and repeat stack, depth, nested serial numbers. Errors on unmatched begin/end. |
| n/a | `R/cli.R` | optparse command line interface `cli()`. |
| 4+ | not built | docx rendering (officer), choice lists and `choice_filter` display, `relevant` (skip logic) display, media embedding, `settings` `form_title`, output path defaulting. |

## Directory layout

- `R/`, `man/`, `NAMESPACE`: package code and generated docs.
- `tests/testthat/`: tests (helpers in `helper-fixtures.R`).
- `inst/extdata/multilang/`: synthetic multilingual example form `multilang.xlsx`, external choice list `enumname.csv`, choice images (apple, samsung, vivo, xiaomi .png), question image `survey.jpg`.
- `dev/`: manual run scripts such as `smoke.R` (Rbuildignored).
- `outputs/`: gitignored scratch outputs (generated .docx).
- `data/raw/`, `data/processed/`: gitignored; see Data handling.
- `.claude/settings.json`: enables the `survey-manager@masud-tools` plugin (has an `xlsform-questionnaire-designer` skill) and denies reads of the data folders.

## Dev commands

devtools is not installed. R 4.6.

```sh
Rscript -e 'roxygen2::roxygenise()'          # docs and NAMESPACE
Rscript -e 'testthat::test_local()'          # tests
Rscript -e 'pkgload::load_all()'             # load (check: cat(requireNamespace("pkgload")))
# build and check from a temp dir so artifacts stay out of the repo
cd "$(mktemp -d)" && R CMD build "$OLDPWD" && R CMD check --no-manual convertxlsform_*.tar.gz
```

CI: GitHub Actions R-CMD-check (check-standard).

## Conventions

- Style: magrittr `%>%`, internal helpers prefixed with `.`, errors with `rlang::abort()`, progress messages with `message()` gated by `verbose`.
- Error messages name the sheet, column, and row number (as in Excel) so the author can fix the XLSForm.
- Keep Imports minimal. Add a dependency only when base R or an existing Import cannot do the job.
- Roxygen with markdown (`Roxygen: list(markdown = TRUE)`). Regenerate docs after changing comments.
- Every new behavior gets a testthat test. Build fixtures in `tests/testthat/helper-fixtures.R` or keep them in `inst/extdata`.
- Prefer editing existing files over creating parallel versions.
- US English, but keep UN terms (for example "programme"). No em-dashes in code, docs, or messages.
- R source is ASCII only. Use `\uXXXX` escapes for non-ASCII characters (as in v0.2.1), required for R CMD check.
- Never fabricate figures, citations, or sources in docs or examples.

## Data handling

XLSForms are instruments, not microdata. Choice lists can still be personal data (for example enumerator or location lists). Those files go in `data/raw/`, are never committed, and Claude does not open them. Real XLSForms or anything sensitive also go in `data/raw/`. The synthetic form in `inst/extdata/multilang` is the only committed example. See `rules/data-handling.md` in the global config.

## Decisions made (2026-10-12)

- Rewrite on branch `rewrite`; keep `main` as v0.2.1.
- Synthetic example form in `inst/extdata/multilang`.
- Development version is 0.2.1.9000.

## Open decisions

- Pipe style: keep `%>%` or migrate to native `|>` (the owner prefers `|>` for new code). (Keep ``%>%).
- Keep v0.2.1 argument names for backward compatibility (`selected_language`, `number_questions`) or use the new ones (`lang`, `nested_serial_numbers`). (Use the new ones).
- Whether and when to submit to CRAN. (Will submit to CRAN).

## Status and next steps

Phases 1 to 3 and the CLI exist. No .docx is written yet.

Known bugs (found 2026-10-12 when the tests were added; the tests encode the intended behavior and fail until fixed):

- `R/read_validate.R` `pick_col()`: matches only `label::<lang>` exactly, so headers like `label::English (en)` (as in the example form) are never found.
- `R/structure.R` `purrr::list_rbind(items)`: items are plain lists, not data frames, so `.build_structure()` always errors.
- `R/structure.R` `begin_*` branch: groups never bump the parent counter, giving serials "0.1" instead of "1.1".
- `R/read_validate.R` `.parse_parameters_vectorized()`: `"[,\\s]+"` without `perl = TRUE` splits on the letter "s".
- `R/read_validate.R` `.parse_base_type()`: leading whitespace returns NA.
- R CMD check: non-ASCII `…` and `–` in messages; `library(optparse)` in `cli()`; unused `tidyr` Import; undeclared globals (`type`, `required`, `parameters`, `n`).

Roadmap:

1. Phase 4: docx skeleton with officer (title from `settings` `form_title`, mandatory-question note, page size and margins).
2. Question rendering: numbering, notes, hints, constraint messages, required asterisk.
3. Answer layouts: select_one / select_multiple bullets, text boxes, range choices, `select_one_from_file`.
4. Choice lists and `choice_filter` display, external CSV choice lists.
5. `relevant` (skip logic) display, controlled by `include_relevants`.
6. Media: question images (max 2 in) and choice images (max 1 in) from `media_root`, warn on missing files.
7. Output path defaulting, then reach v0.2.1 parity, open PR to `main`, bump to 0.3.0.
8. Settle open decisions above before the PR.
