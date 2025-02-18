# convertxlsform

**convertxlsform** is an R package that converts [XLSForm questionnaires](https://xlsform.org/en/)—commonly used in [Open Data Kit (ODK)](https://opendatakit.org/software/) surveys—into neatly formatted Microsoft Word documents. This package was tested using [KoboToolbox](https://www.kobotoolbox.org/), but should ideally work for most XLSForm based survey platforms. The package supports multiple languages (if language-specific columns are present) and applies several formatting options, such as:

- Prefixing required questions with an asterisk (*)
- Formatting group/repeat sections with clear headers
- Rendering hints in italics
- Enclosing constraint messages in parentheses and rendering them in italics

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)

## Overview

In many survey projects, questionnaires are built using XLSForms (typically edited with tools like KoboToolbox or ODK Build). However, when reviewing across different teams, survey designers need a formatted version of the questionnaire. **convertxlsform** automates this conversion by:

- Reading the XLSForm’s *settings*, *survey*, and *choices* sheets
- Extracting and formatting question names, types, labels, hints, and constraints
- Respecting language-specific columns if present (e.g., `label::English (en)`)
- Producing a well-formatted Word document (.docx) for use in PAPI data collection

## Features

- **Multi-language support:** If your XLSForm has language-specific columns (e.g., `label::en`, `hint::en`), the package uses these columns. Otherwise, it defaults to the base columns and omits the language header.
- **Required question marking:** Questions marked as required (with a `TRUE` value in the required column) have an asterisk (*) prefixed to the question name.
- **Custom formatting for groups and repeats:** Group and repeat sections are clearly marked with headers that include relevance and calculation information (if available).
- **Answer area placeholders:** Depending on the question type (text, integer, geopoint, calculate, select_one/select_multiple), the package prints an appropriate answer placeholder or calculates the result.
- **HTML formatting support:** Currently not supported, but is part of the roadmap.

## Installation

You can install **convertxlsform** directly from GitHub using the `devtools` package. First, make sure you have the `devtools` package installed:

```r
install.packages("devtools")
```

Then install convertxlsform:
```r
devtools::install_github("masud90/convertxlsform")
```

## Usage
After installation, load the package and call the main function `convertxlsform()` with the path to your XLSForm. For example:

```r
# Load the package
library(convertxlsform)

# Convert your XLSForm to a formatted DOCX.
convertxlsform("path/to/your/form.xlsx", selected_language = "English (en)")
```
### Function Arguments
- `xlsform_path`: A character string specifying the path to the XLSForm (.xlsx file).
- `selected_language`: A character string indicating the language to use for label and hint columns (default is "en"). If the XLSForm does not have language-specific columns, the package will default to using the base columns (e.g.: label, hint, constraint etc.).
- `output_docx`: (Optional) A character string specifying the output DOCX filename. If NULL (the default), the function uses the XLSForm’s base name appended with "_PAPER.docx".

## Examples
Below is an example workflow:

```r
# Load the package
library(convertxlsform)
# Convert an XLSForm (e.g., "survey.xlsx") to a Word document.
# This will generate a file named "survey_PAPER.docx" in the working directory.
convertxlsform("survey.xlsx", selected_language = "English (en)")
```

The generated Word document will include:

- A title (from the XLSForm settings)
- A "Selected language:" header (if language-specific columns exist)
- Each question formatted with its type, relevance (if provided), and other details
- Required questions marked with an asterisk (*)
- Constraint messages (if any) in italics and enclosed in parentheses
- Answer area placeholders for questions without specified answer choices (text, integer, geopoint, etc.)

## Contributing
Contributions are welcome! If you find issues, have feature requests, or wish to improve the code, please:

- Fork the repository.
- Create a new branch with your changes.
- Submit a pull request detailing your changes.
- Please ensure that your code adheres to our coding style and passes devtools::check() before submitting your pull request.

## License
This project is licensed under the MIT License. See the [license](LICENSE) file for details.

