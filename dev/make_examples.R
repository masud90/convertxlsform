# Generates the synthetic "all types" example form in inst/extdata/all_types/.
#
# Run from the project root:
#   Rscript dev/make_examples.R
#
# Everything here is fictional: a household survey on water, food and assets
# in District A, B and C. No personal data. Output is deterministic.
# Requires the writexl and utils packages (utils ships with R).
# Non-ASCII text is written with \u escapes so this file stays ASCII.

if (!file.exists("DESCRIPTION")) {
  stop("Run this script from the project root: Rscript dev/make_examples.R",
       call. = FALSE)
}
if (!requireNamespace("writexl", quietly = TRUE)) {
  stop("Package 'writexl' is needed to build the example.", call. = FALSE)
}

out_dir <- file.path("inst", "extdata", "all_types")
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# ---------------------------------------------------------------------------
# survey sheet
# ---------------------------------------------------------------------------

survey_cols <- c(
  "type", "name",
  "label::English (en)", "label::French (fr)",
  "hint::English (en)", "hint::French (fr)",
  "required", "required_message::English (en)",
  "relevant", "constraint", "constraint_message::English (en)",
  "appearance", "choice_filter", "repeat_count", "parameters",
  "calculation", "default", "media::image::English (en)"
)

srow <- function(type, name = NA_character_,
                 label = NA_character_, label_fr = NA_character_,
                 hint = NA_character_, hint_fr = NA_character_,
                 required = NA_character_, required_message = NA_character_,
                 relevant = NA_character_, constraint = NA_character_,
                 constraint_message = NA_character_,
                 appearance = NA_character_, choice_filter = NA_character_,
                 repeat_count = NA_character_, parameters = NA_character_,
                 calculation = NA_character_, default = NA_character_,
                 image = NA_character_) {
  vals <- list(type, name, label, label_fr, hint, hint_fr, required,
               required_message, relevant, constraint, constraint_message,
               appearance, choice_filter, repeat_count, parameters,
               calculation, default, image)
  stopifnot(length(vals) == length(survey_cols))
  out <- as.data.frame(lapply(vals, as.character), stringsAsFactors = FALSE)
  names(out) <- survey_cols
  out
}

survey_rows <- list(
  # metadata
  srow("start", "start"),
  srow("end", "end"),
  srow("today", "today"),
  srow("deviceid", "deviceid"),

  # introduction with an image
  srow("note", "intro_note",
       label = "Household survey on water, food and assets in District A, B and C. Read this introduction to the respondent before starting.",
       label_fr = "Enqu\u00eate sur l'eau, l'alimentation et les biens des m\u00e9nages dans les districts A, B et C. Lisez cette introduction au r\u00e9pondant avant de commencer.",
       image = "household.png"),

  # consent and a follow-up that depends on it
  srow("select_one yesno", "consent",
       label = "Does the respondent agree to take part in the survey?",
       label_fr = "Le r\u00e9pondant accepte-t-il de participer \u00e0 l'enqu\u00eate ?",
       hint = "Read the consent statement aloud first.",
       required = "yes", required_message = "An answer is required"),
  srow("text", "follow_up",
       label = "Enumerator code for this interview",
       label_fr = "Code de l'enqu\u00eateur pour cet entretien",
       required = "${consent} = 'yes'"),

  # Household group
  srow("begin_group", "hh",
       label = "Household",
       label_fr = "M\u00e9nage"),
  srow("text", "hh_code",
       label = "Household code",
       label_fr = "Code du m\u00e9nage",
       hint = "Use the code printed on the household card.",
       required = "yes", required_message = "Household code is required"),
  srow("integer", "hh_size",
       label = "Number of household members",
       label_fr = "Nombre de membres du m\u00e9nage",
       required = "yes",
       constraint = ". >= 0 and . <= 30",
       constraint_message = "Enter a whole number from 0 to 30"),
  srow("decimal", "land_ha",
       label = "Area of land cultivated (hectares)",
       label_fr = "Superficie cultiv\u00e9e (hectares)",
       hint = "Enter 0 if no land is cultivated.",
       hint_fr = "Saisir 0 si aucune terre n'est cultiv\u00e9e."),
  srow("date", "visit_date",
       label = "Date of visit",
       label_fr = "Date de la visite"),
  srow("time", "visit_time",
       label = "Time of visit",
       label_fr = "Heure de la visite"),
  srow("dateTime", "last_rain",
       label = "Date and time of the last rainfall",
       label_fr = "Date et heure de la derni\u00e8re pluie"),
  srow("select_one water_source or_other", "water_source",
       label = "Main source of drinking water",
       label_fr = "Principale source d'eau potable",
       hint = "Choose one answer.",
       hint_fr = NA),
  srow("select_multiple assets", "assets",
       label = "Household assets",
       label_fr = "Biens du m\u00e9nage",
       hint = "Select all that apply.",
       hint_fr = "Cocher toutes les r\u00e9ponses applicables."),
  srow("rank priorities", "priorities",
       label = "Rank the household's priorities from first to last",
       label_fr = "Classer les priorit\u00e9s du m\u00e9nage de la premi\u00e8re \u00e0 la derni\u00e8re"),
  srow("range", "water_hours",
       label = "Hours spent collecting water per day",
       label_fr = "Heures pass\u00e9es \u00e0 chercher de l'eau par jour",
       parameters = "start=0 end=10 step=1"),
  srow("range", "food_share",
       label = "Share of household income spent on food (%)",
       label_fr = "Part du revenu du m\u00e9nage d\u00e9pens\u00e9e en nourriture (%)",
       parameters = "start=0 end=100 step=5"),
  srow("acknowledge", "ack_understood",
       label = "I confirm that the respondent understood the purpose of the survey",
       label_fr = "Je confirme que le r\u00e9pondant a compris l'objectif de l'enqu\u00eate"),
  srow("barcode", "asset_tag",
       label = "Scan the household asset tag",
       label_fr = "Scanner l'\u00e9tiquette du bien du m\u00e9nage"),

  # nested labelled group inside Household
  srow("begin_group", "water_food",
       label = "Water and food",
       label_fr = "Eau et alimentation"),
  srow("select_multiple food_sources", "food_sources",
       label = "Main sources of food in the past week",
       label_fr = "Principales sources d'alimentation de la derni\u00e8re semaine"),
  srow("text", "food_note",
       label = "Reason for not buying food at the market",
       label_fr = "Raison de ne pas acheter de la nourriture au march\u00e9",
       relevant = "not(selected(${food_sources}, 'market'))"),
  srow("integer", "radio_hours",
       label = "Hours per week the radio is used",
       label_fr = "Heures par semaine d'utilisation de la radio",
       relevant = "${hh_size} > 2 and selected(${assets}, 'radio')"),
  srow("end_group"),

  # unlabelled field-list group (transparent)
  srow("begin_group", "details",
       appearance = "field-list"),
  srow("decimal", "asset_value",
       label = "Estimated value of household assets (USD)",
       label_fr = "Valeur estim\u00e9e des biens du m\u00e9nage (USD)",
       relevant = "count-selected(${assets}) > 1"),
  srow("select_one yesno", "owns_livestock",
       label = "Does the household own livestock?",
       label_fr = "Le m\u00e9nage poss\u00e8de-t-il du b\u00e9tail ?"),
  srow("end_group"),

  srow("integer", "livestock_count",
       label = "Number of livestock animals",
       label_fr = "Nombre d'animaux d'\u00e9levage",
       relevant = "${owns_livestock} = 'yes'"),
  srow("end_group"),

  # repeat group driven by hh_size
  srow("begin_repeat", "members",
       label = "Household member",
       label_fr = "Membre du m\u00e9nage",
       repeat_count = "${hh_size}"),
  srow("text", "member_name",
       label = "First name of member",
       label_fr = "Pr\u00e9nom du membre",
       required = "yes"),
  srow("integer", "member_age",
       label = "Age of member (years)",
       label_fr = "\u00c2ge du membre (ann\u00e9es)",
       constraint = ". >= 0 and . <= 120",
       constraint_message = "Enter an age from 0 to 120"),
  srow("select_one sex", "member_sex",
       label = "Sex of member",
       label_fr = "Sexe du membre"),
  srow("end_repeat"),

  # geo, media
  srow("geopoint", "location",
       label = "Location of the household (GPS)",
       label_fr = "Localisation du m\u00e9nage (GPS)",
       image = "map.jpg"),
  srow("geotrace", "trace_route",
       label = "Walking route to the water point",
       label_fr = "Itin\u00e9raire \u00e0 pied jusqu'au point d'eau"),
  srow("geoshape", "plot_shape",
       label = "Boundary of the cultivated plot",
       label_fr = "Limite de la parcelle cultiv\u00e9e"),
  srow("image", "photo_dwelling",
       label = "Photo of the dwelling",
       label_fr = "Photo du logement"),
  srow("audio", "audio_comment",
       label = "Audio recording of the respondent's comments",
       label_fr = "Enregistrement audio des commentaires du r\u00e9pondant"),
  srow("video", "video_water",
       label = "Video of the water point",
       label_fr = "Vid\u00e9o du point d'eau"),
  srow("file", "consent_form",
       label = "Scanned consent form (PDF)",
       label_fr = "Formulaire de consentement scann\u00e9 (PDF)"),

  # cascading selects
  srow("select_one district", "district",
       label = "District",
       label_fr = "District"),
  srow("select_one village", "village",
       label = "Village (filtered by district)",
       label_fr = "Village (filtr\u00e9 par district)",
       choice_filter = "district = ${district}"),

  # choices from external files (missing_list.csv is intentionally absent)
  srow("select_one_from_file villages.csv", "village_file",
       label = "Village (from file)",
       label_fr = "Village (depuis un fichier)"),
  srow("select_multiple_from_file missing_list.csv", "missing_items",
       label = "Items from an external list",
       label_fr = "\u00c9l\u00e9ments d'une liste externe"),

  # long list
  srow("select_one crop", "main_crop",
       label = "Main crop",
       label_fr = "Culture principale"),

  # calculation, required expression, multiline text, closing note
  srow("calculate", "hh_size_double",
       calculation = "${hh_size} * 2"),
  srow("text", "comments",
       label = "Additional comments",
       label_fr = "Commentaires suppl\u00e9mentaires",
       appearance = "multiline"),
  srow("note", "closing",
       label = "Thank you. The survey is complete.",
       label_fr = "Merci. L'enqu\u00eate est termin\u00e9e.")
)

survey <- do.call(rbind, survey_rows)

# ---------------------------------------------------------------------------
# choices sheet
# ---------------------------------------------------------------------------

choices_cols <- c(
  "list_name", "name",
  "label::English (en)", "label::French (fr)",
  "media::image::English (en)", "district"
)

crow <- function(list_name, name, label, label_fr = NA_character_,
                 image = NA_character_, district = NA_character_) {
  out <- as.data.frame(
    lapply(list(list_name, name, label, label_fr, image, district),
           as.character),
    stringsAsFactors = FALSE
  )
  names(out) <- choices_cols
  out
}

yesno <- list(
  crow("yesno", "yes", "Yes", "Oui"),
  crow("yesno", "no", "No", "Non")
)

water <- list(
  crow("water_source", "tap", "Tap", "Robinet", image = "tap.png"),
  crow("water_source", "well", "Well", "Puits", image = "well.png"),
  crow("water_source", "river", "River", "Rivi\u00e8re", image = "river.png"),
  crow("water_source", "spring", "Spring", "Source")
)

assets <- list(
  crow("assets", "radio", "Radio", "Radio"),
  crow("assets", "phone", "Mobile phone", "T\u00e9l\u00e9phone mobile"),
  crow("assets", "bicycle", "Bicycle", "Bicyclette"),
  crow("assets", "livestock", "Livestock", "B\u00e9tail")
)

priorities <- list(
  crow("priorities", "water", "Water", "Eau"),
  crow("priorities", "food", "Food", "Nourriture"),
  crow("priorities", "shelter", "Shelter", "Abri"),
  crow("priorities", "health", "Health", "Sant\u00e9")
)

sex <- list(
  crow("sex", "male", "Male", "Homme"),
  crow("sex", "female", "Female", "Femme")
)

districts <- list(
  crow("district", "district_a", "District A", "District A"),
  crow("district", "district_b", "District B", "District B"),
  crow("district", "district_c", "District C", "District C")
)

# village rows carry the district filter column
villages <- list(
  crow("village", "village_a1", "Village A1", "Village A1", district = "district_a"),
  crow("village", "village_a2", "Village A2", "Village A2", district = "district_a"),
  crow("village", "village_b1", "Village B1", "Village B1", district = "district_b"),
  crow("village", "village_b2", "Village B2", "Village B2", district = "district_b"),
  crow("village", "village_c1", "Village C1", "Village C1", district = "district_c"),
  crow("village", "village_c2", "Village C2", "Village C2", district = "district_c")
)

food_sources <- list(
  crow("food_sources", "market", "Market", "March\u00e9"),
  crow("food_sources", "farm", "Own farm", "Propre ferme"),
  crow("food_sources", "aid", "Aid distribution", "Distribution d'aide"),
  crow("food_sources", "gathering", "Gathering wild food", "Cueillette")
)

crops <- lapply(seq_len(30), function(i) {
  crow("crop",
       sprintf("crop_%02d", i),
       sprintf("Crop %d", i),
       sprintf("Culture %d", i))
})

choices <- do.call(rbind, c(yesno, water, assets, priorities, sex,
                            districts, villages, food_sources, crops))

# ---------------------------------------------------------------------------
# settings sheet
# ---------------------------------------------------------------------------

settings <- data.frame(
  form_title = "All question types example",
  form_id = "all_types",
  version = "2026101201",
  default_language = "English (en)",
  stringsAsFactors = FALSE
)

# ---------------------------------------------------------------------------
# workbook
# ---------------------------------------------------------------------------

xlsx_path <- file.path(out_dir, "all_types.xlsx")
writexl::write_xlsx(
  list(survey = survey, choices = choices, settings = settings),
  path = xlsx_path
)
cat("Wrote", xlsx_path, "\n")

# ---------------------------------------------------------------------------
# villages.csv (choices from file)
# ---------------------------------------------------------------------------

village_file <- data.frame(
  name = sprintf("village_%d", 1:5),
  `label::English (en)` = sprintf("Village %d", 1:5),
  `label::French (fr)` = sprintf("Village %d", 1:5),
  check.names = FALSE,
  stringsAsFactors = FALSE
)
csv_path <- file.path(out_dir, "villages.csv")
utils::write.csv(village_file, csv_path, row.names = FALSE,
                 fileEncoding = "UTF-8")
cat("Wrote", csv_path, "\n")

# ---------------------------------------------------------------------------
# media: small base-graphics pictures, 200 x 150 px
# ---------------------------------------------------------------------------

draw_picture <- function(shape, label, bg, fg) {
  plot.new()
  plot.window(xlim = c(0, 200), ylim = c(0, 150), xaxs = "i", yaxs = "i")
  rect(0, 0, 200, 150, col = bg, border = NA)
  switch(shape,
    house = {
      polygon(c(50, 100, 150), c(70, 115, 70), col = fg, border = NA)
      rect(60, 30, 140, 70, col = fg, border = NA)
      rect(92, 35, 108, 70, col = bg, border = NA)
    },
    tap = {
      rect(40, 40, 60, 110, col = fg, border = NA)
      rect(60, 95, 140, 110, col = fg, border = NA)
      rect(130, 60, 145, 110, col = fg, border = NA)
      rect(30, 20, 70, 40, col = fg, border = NA)
    },
    well = {
      symbols(100, 80, circles = 45, inches = FALSE, add = TRUE,
              bg = fg, fg = fg)
      symbols(100, 80, circles = 30, inches = FALSE, add = TRUE,
              bg = bg, fg = bg)
      rect(70, 125, 130, 140, col = fg, border = NA)
    },
    river = {
      for (y in c(45, 75, 105)) {
        xs <- seq(0, 200, length.out = 60)
        lines(xs, y + 8 * sin(xs / 12), col = fg, lwd = 6)
      }
    },
    map = {
      for (x in seq(20, 180, by = 40)) lines(c(x, x), c(15, 135), col = fg)
      for (y in seq(15, 135, by = 30)) lines(c(10, 190), c(y, y), col = fg)
      symbols(120, 70, circles = 10, inches = FALSE, add = TRUE,
              bg = "red", fg = "red")
    }
  )
  text(100, 15, label, col = fg, cex = 1.1)
  invisible(NULL)
}

write_png <- function(file, shape, label, bg = "white", fg = "#1f4e79") {
  png(file, width = 200, height = 150, units = "px", bg = bg)
  on.exit(dev.off())
  par(mar = c(0, 0, 0, 0))
  draw_picture(shape, label, bg, fg)
  cat("Wrote", file, "\n")
}

write_jpeg <- function(file, shape, label, bg = "white", fg = "#1f4e79") {
  jpeg(file, width = 200, height = 150, units = "px", quality = 90, bg = bg)
  on.exit(dev.off())
  par(mar = c(0, 0, 0, 0))
  draw_picture(shape, label, bg, fg)
  cat("Wrote", file, "\n")
}

write_png(file.path(out_dir, "household.png"), "house", "Household")
write_png(file.path(out_dir, "tap.png"), "tap", "Tap")
write_png(file.path(out_dir, "well.png"), "well", "Well")
write_png(file.path(out_dir, "river.png"), "river", "River")
write_jpeg(file.path(out_dir, "map.jpg"), "map", "Sketch map")

cat("Done.\n")
