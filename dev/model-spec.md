# Internal contract

Every worker builds against this file. Change it only through Opus review. Base R data frames (not tibbles), `%>%` allowed, errors via `rlang::abort()`, user messages via `cli::cli_inform()` / `cli::cli_warn()` gated by `verbose`. R source ASCII only (`\uXXXX` escapes). Internal functions are prefixed with `.` and documented with `@noRd`.

## 1. Parsed form: `.read_xlsform(path, lang, media_root, verbose)` returns class `cxf_form`

A list:

| Element | Type | Contents |
|---|---|---|
| `path` | chr(1) | normalized input path |
| `media_root` | chr(1) | normalized; default `dirname(path)` |
| `settings` | named list | lower-cased keys from the first data row of `settings` (or empty list); always has `form_title`, `form_id`, `version`, `default_language` (NA if absent) |
| `languages` | data.frame | `suffix` (text after `::`, e.g. `English (en)`), `code` (e.g. `en`, NA if none in parentheses and suffix is not a 2-3 letter code), `name` (e.g. `English`) |
| `lang` | list | `requested` (chr or NA), `suffix` (chr or NA when single-language), `mode` one of `"multilingual"`, `"single"`, `"fallback"` |
| `survey` | data.frame | see below, one row per non-blank survey row, in sheet order |
| `choices` | data.frame | `list_name`, `name`, `label`, `image`, `row`, plus every other original column kept verbatim (filter columns for `choice_filter`) |

### `survey` columns

| Column | Type | Notes |
|---|---|---|
| `row` | int | Excel row number (header is row 1) |
| `type` | chr | raw type, trimmed, internal whitespace collapsed |
| `base_type` | chr | lower-case, normalized: `begin group` -> `begin_group`, `dateTime` -> `datetime`, `select one` -> `select_one` |
| `type_arg` | chr | list name or file name after the base type (`select_one yesno` -> `yesno`; `select_one_from_file x.csv` -> `x.csv`), NA otherwise |
| `or_other` | lgl | TRUE when type ends with ` or_other` |
| `kind` | chr | display kind from the type registry (section 2) |
| `name` | chr | |
| `label`, `hint`, `guidance_hint`, `required_message`, `constraint_message`, `image`, `audio`, `video` | chr | resolved for the chosen language (section 3); `media::image` treated as `image` |
| `required` | lgl | `yes`/`true`/`1`/`TRUE()` -> TRUE; anything else non-empty (an expression) -> TRUE with `required_expr` set |
| `required_expr` | chr | NA unless `required` holds an expression |
| `relevant`, `constraint`, `choice_filter`, `repeat_count`, `calculation`, `default`, `appearance`, `parameters`, `read_only` | chr | NA when absent |
| `params` | list | named list parsed from `parameters` (`start=1 end=10 step=1`, separators: spaces, commas, semicolons); numeric values coerced |
| `depth` | int | number of enclosing groups/repeats (top level = 0); a `begin_*` row has the depth of its parent; its `end_*` row matches |
| `parent` | int | index (in `survey`) of the enclosing `begin_*` row, NA at top level |
| `serial` | chr | display number, NA when unnumbered (section 4) |

## 2. Type registry (`R/types.R`)

`.type_registry()` returns a data.frame `base_type`, `kind`, `subtype`. Kinds:

| kind | base types |
|---|---|
| `begin_group`, `end_group`, `begin_repeat`, `end_repeat` | same |
| `select_one` | select_one, select_one_from_file, select_one_external |
| `select_multiple` | select_multiple, select_multiple_from_file |
| `rank` | rank |
| `text` | text |
| `number` | integer (subtype `integer`), decimal (subtype `decimal`) |
| `range` | range |
| `date` / `time` / `datetime` | date / time / datetime |
| `geo` | geopoint, geotrace, geoshape (subtype = base type) |
| `media` | image, audio, video, file, background-audio (subtype = base type) |
| `barcode` | barcode |
| `acknowledge` | acknowledge, trigger |
| `note` | note |
| `hidden` | calculate, hidden, xml-external, csv-external, start, end, today, deviceid, subscriberid, simserial, phonenumber, username, email, audit, start-geopoint |
| `unknown` | anything else (warning citing the row; rendered like `text`) |

## 3. Language resolution (`R/language.R`)

- Languages come from every column matching `^(label|hint|guidance_hint|required_message|constraint_message|image|media::image|audio|video)::(.+)$` in survey and choices. Suffix parsing: `English (en)` -> name `English`, code `en`; `en` -> code `en`, name NA; `English` -> name `English`, code NA.
- `lang` argument matches, case-insensitively, a suffix exactly, a code, or a name. `"English (en)"`, `"en"` and `"English"` all select suffix `English (en)`.
- `lang = NULL`: use `settings$default_language` if it matches; else the first language in column order; else single-language mode.
- Requested language not found: warn, then use default/first language (`mode = "fallback"`).
- Per field, the value is taken from `field::<suffix>`; if that column is absent or the cell is empty, fall back to the plain `field` column, then to the default language's column. A plain `label` column with no `::` columns is `mode = "single"`.
- `form_title::<suffix>` in settings is honoured the same way.

## 4. Structure and numbering (`R/structure.R`)

- Stack-based pass over `kind`. Errors (abort) cite Excel rows: `end_group` with no open group, `end_repeat` closing a group (and vice versa), unclosed `begin_*` at end of sheet.
- Numbered items: `begin_group`, `begin_repeat`, and every question kind except `note`, `hidden`, and the `end_*` kinds.
- `nested_serial_numbers = TRUE`: counters per depth. A group at depth d takes the next counter at depth d; its children are numbered `<group serial>.<k>`. Example: q (1), group (2) {q (2.1), group (2.2) {q (2.2.1)}}, q (3).
- Groups with no label (or `appearance` containing `field-list` and no label) are transparent: not numbered, no heading, children numbered at the parent level as if the group were absent.
- `nested_serial_numbers = FALSE`: groups unnumbered; questions numbered 1..n across the whole form.
- Clarified after M1 (2026-10-12): only labelless `begin_group` rows are transparent; a labelless repeat is still numbered and gets a banner. `depth` counts every enclosing group including transparent ones. The model derives `level` (heading level, 1 to 3, capped) from visible enclosing sections only. `.read_xlsform()` numbers with `nested_serial_numbers = TRUE`; callers renumber with `.build_structure(form$survey, FALSE)`. `cxf_form` also carries `sheets` and `found` (used by validation).

## 5. Expression humanizing (`R/expressions.R`)

`.humanize_expr(expr, form, field)` returns chr(1) (or NA for NA input). Rules applied in order, on a best-effort tokenizer; when anything is not recognized, the original expression is returned with only `${x}` substitution done.

- `${x}` -> `Q<serial> (x)` when x has a serial, else `x`.
- `selected(${x}, 'a')` -> `Q2.1 (x) includes "Label of a"` (label from x's choice list, falls back to `'a'`).
- `not(selected(...))` -> `... does not include ...`.
- `${x} = 'a'` where x is a select_one -> `Q2.1 (x) is "Label of a"`; `!=` -> `is not`.
- `count-selected(${x})` -> `number selected in Q.. (x)`.
- Operators: `=` -> `=`, `!=` -> `is not`, `>=` `<=` `>` `<` kept, ` and ` / ` or ` kept, `.` alone (constraint) -> `the answer`, `''` -> `empty`, `today()` -> `today`.
- `field` is `"relevant"`, `"constraint"`, `"choice_filter"`, or `"repeat_count"`; it only changes the `.` handling.

## 6. Document model (`R/model.R`)

`.build_model(form, opts)` returns class `cxf_model`: list with `meta` and `blocks`.

`meta`: `title`, `form_id`, `version`, `language` (display name or NA), `generated` (Date), `legend` (chr vector).

`blocks`: list, in order. Every block is a list with `block` (chr), `row`, `depth`, `serial`, `name`. Block types:

- `section`: `label`, `hint`, `is_repeat` (lgl), `repeat_text` (humanized repeat_count or NA), `relevant_text`, `image` (media record or NULL). Emitted for numbered/labelled `begin_*` rows only.
- `section_end`: emitted for `end_repeat` only (closes the repeat banner).
- `note`: `label`, `hint`, `image`, `relevant_text`.
- `question`: `label`, `hint`, `required` (lgl), `kind`, `subtype`, `relevant_text`, `constraint_text` (humanized constraint, with `constraint_message` appended in parentheses when present), `choice_filter_text`, `image` (media record or NULL), `answer` (below), `calculation` (only for hidden kind when `include_hidden`).

Media record: `list(path = <abs path or NA>, file = <as written>, found = lgl, width_in, height_in)`; size scaled to `max_height_in` (2 for questions, 0.6 for choices) and to page text width.

`answer` by kind:

| kind | answer fields |
|---|---|
| select_one / select_multiple / rank | `choices` data.frame (`name`, `label`, `image` list-column of media records or NULL), `symbol` (`"circle"`, `"square"`, `"rank"`), `or_other` lgl, `n_total`, `truncated` lgl, `source` (`"choices"` or file name), `source_found` lgl |
| text | `lines` (4 if appearance has `multiline`, else 2) |
| number | `decimal` lgl, `range_text` (from constraint when simple, else NA) |
| range | `start`, `end`, `step`, `values` (numeric vector when <= 20 values, else NULL) |
| date / time / datetime | `pattern` (`"__ / __ / ____  (dd/mm/yyyy)"`, `"__ : __  (hh:mm)"`, both for datetime) |
| geo | `subtype`, `fields` (geopoint: Latitude, Longitude, Altitude, Accuracy; geotrace/geoshape: "List of points" lines) |
| media | `subtype`, `prompt` (e.g. "[Photo to be taken]") |
| barcode | `lines` = 1 |
| acknowledge | `prompt` = "OK" tick |
| hidden | `calculation` |
| unknown | `lines` = 2 |

Choices from file (`*_from_file`): read `media_root/<file>` if it exists (`.csv` via `utils::read.csv(encoding = "UTF-8")`, `.xml` parsed in base R, one row per `<item>` and one column per child element); label column chosen with the same language rules; if missing, `source_found = FALSE` and the renderer prints "Choices from <file>".
`max_choices`: when a list is longer, keep the first `max_choices` and set `truncated = TRUE`.

## 7. `xlsform_outline()` (exported)

Data frame, one row per block: `row`, `serial`, `depth`, `block`, `kind`, `name`, `label`, `required`, `n_choices`, `relevant_text`, `constraint_text`. Used for snapshots and by users to inspect a form.

## 8. Renderer (`R/render_docx.R`, `R/answer_layouts.R`)

`.render_docx(model, opts, path)` writes the file and returns `path`. Uses `officer::read_docx(template)` with the shipped template `system.file("templates", "convertxlsform.docx", package = "convertxlsform")` unless `opts$template` is set. Styles referenced by name: `Title`, `heading 1`, `heading 2`, `heading 3`, `Normal`, plus custom paragraph styles in the template: `cxf Question`, `cxf Hint`, `cxf Logic`, `cxf Choice`, `cxf Note`. Fallback when a style is missing from a user template: `Normal` with direct formatting.

Symbols (`R/utils.R`, written as escapes in R source): circle U+25CB, square U+25A1, rank box U+25A1, tick box U+25A1. Use only WGL4 glyphs present in Arial and Liberation Sans: U+25EF and U+2610 fall back to other fonts in LibreOffice, and U+25EF renders filled there. Symbol runs use the document font at `font_size + 3` pt. Also add `cxf Banner` to the custom styles list above.

Heading styles are referenced by style name (`"heading 1"`), never by id: officer's default template has localized ids (`Titre1`).

Character styles `cxf Required` and `cxf Code` exist in the template for users, but officer cannot apply a character style to a run. The renderer formats runs directly: required asterisk `fp_text(bold = TRUE, color = "#C00000")`, codes and variable names `fp_text(font.size = font_size - 2, color = "#595959")`.

## 9. Options object

`.cxf_options(...)` validates and returns a list with every option named in `dev/plan.md` (Public API), plus derived `text_width_in` from page size, orientation and margins.

## 10. Revisions after the Fable design review (2026-10-12)

These override earlier sections where they conflict.

Model (M2b):

1. `${var}` inside label, hint, guidance_hint and note text becomes `[Q2.2.1]` when the target is numbered, else `[name]`.
2. Logic display. `constraint_text`: if `constraint_message` exists show only the message; else the humanized expression; if humanizing fell back to verbatim and `show_variable_names = FALSE`, drop it (NA). `relevant_text`: question references render as `Q3.16` when `show_variable_names = FALSE`, `Q3.16 (owns_livestock)` when TRUE; verbatim fallbacks are kept (skip logic must not silently vanish) but only with `${}` substitution.
3. Labels, hints and notes may contain line breaks (`\r\n`, `\n`); normalize to `\n`. The renderer splits on `\n` and uses `officer::run_linebreak()`.
4. Question blocks carry `appearance` and `guidance_hint` (shown when `include_hints`).
5. Select answers carry `layout`: `"list"` (<= 7 choices), `"columns"` (8 to `max_choices`, with `ncol` 2 or 3: 3 when every label is <= 25 characters, else 2), `"code_box"` (more than `max_choices`, or appearance contains `minimal`/`autocomplete` with more than 7 choices): a write-in "Code: ____" box plus a compact multi-column code table of all choices. `max_choices` default changes from `Inf` to 30 and no longer truncates.
6. `or_other = TRUE` appends a choice `name = "other"`, `label = <strings$other_specify>`, `write_in = TRUE` (every choice row has a `write_in` column, FALSE otherwise).
7. Legend adds a rank line when the form has rank questions.
8. String table: `.cxf_strings(lang_code, overrides)` returns ~15 connective strings (ask_if, constraint, includes, does_not_include, is, is_not, is_empty, is_not_empty, the_answer, number_selected_in, today, repeat_for_each, instance_no, other_specify, legend lines, not_on_paper, reference_no, code). Built in: `en`, `fr`. Other languages fall back to `en`; users override with a new option `strings = NULL` (named list). The humanizer uses these strings.
9. Answer details: geopoint fields are Latitude, Longitude, Accuracy (m) (no altitude). Geotrace, geoshape and media kinds get `prompt = strings$not_on_paper` plus a "Reference no." write-in. Rank choices render with an empty box per choice. Range with <= 20 values: row of circled values; else a box with "start to end, steps of step".
10. Repeats: `section` blocks for repeats carry `repeat_count_value` (numeric when the count is a constant, else NA). New option `repeat_copies = 1L`: the renderer repeats the block content that many times, each with "Instance no. ____".
11. Fonts: new option `font_family_cs = NULL`. When NULL, detect script from the chosen language's labels: Bengali (U+0980 to U+09FF) -> "Noto Sans Bengali"; Arabic (U+0600 to U+06FF) -> "Noto Naskh Arabic"; Devanagari -> "Noto Sans Devanagari"; Ethiopic -> "Noto Sans Ethiopic"; else `font_family`. Stored as `meta$font_family_cs`; `meta$rtl` TRUE for Arabic or Hebrew scripts (RTL layout itself is unsupported in 0.3.0, documented).

Renderer (M3):

12. Selects as borderless 3-column flextables (symbol, code in muted smaller text, label), codes shown when `include_choice_name`; choice images (from v0.2.1 parity) under the label in the same cell, max 0.6 in.
13. Text, number, barcode: bordered single-cell flextables with fixed row heights (0.3 in per line), number box with range_text beside it.
14. Page header with title and version; footer with form_id, generation date and "Page X of Y" (`run_word_field("PAGE")`, `run_word_field("NUMPAGES")`). Table rows must not split across pages; question, hint and logic paragraphs keep with next.
15. `cs.family` set on every text run (`fp_text(font.family, cs.family = meta$font_family_cs)`), complex-script size `font_size + 1`.
16. Template: drop the `cxf Required` and `cxf Code` character styles.

Done in 0.4.0: forward skips ("If No, go to Q5"; `R/skips.R`), roster tables for repeats and grids for `table-list` and `list-nolabel` (`R/render_tables.R`), bilingual stacked output (`lang2`), RTL layout. Still deferred: GeoJSON choice files, annex code lists, rendering of read_only calculations and defaults.
