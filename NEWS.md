# convertxlsform News

## 0.2.1 (2025-07-XX)

- Title field capitalization corrected in DESCRIPTION.
- Added runnable examples in documentation.
- Version bump for CRAN resubmission.

## 0.2.0 (2025-07-19)

**First CRAN release.**

### Highlights

-   **Full XLSForm support**
    -   Reads `survey`, `choices`, and `settings` sheets; errors if any are missing or malformed.\
    -   Preserves HTML/Markdown in `label`, `hint`, and `constraint_message` via `commonmark` → `officer::body_add_html`.
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

### Packaging & internals

-   **Dependencies**: `readxl`, `officer`, `commonmark`, `flextable`, `magick` declared in Imports.\
-   **ASCII compliance**: all Unicode in code replaced with `\uXXXX` escapes.\
-   **Roxygen imports**: explicit `@importFrom` for all used functions.\
-   **Tests**: initial `testthat` scaffold with basic validation tests.

------------------------------------------------------------------------

Thank you for using **convertxlsform**—feedback and issues are welcome on GitHub (<https://github.com/masud90/convertxlsform>)!
