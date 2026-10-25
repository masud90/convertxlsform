# Connective strings used by the humanizer, the model and the renderer, with
# built-in tables for English, French, Spanish, Portuguese, Arabic, Bengali and
# Swahili. ASCII-only source: non-ASCII characters are \u escapes. The
# Spanish, Portuguese, Arabic, Bengali and Swahili tables were machine-drafted
# and await review by native speakers (issue #12).

#' Built-in string table for English
#' @noRd
.strings_en <- function() {
  list(
    ask_if = "Ask if",
    constraint = "Constraint",
    choice_filter = "Choices shown if",
    includes = "includes",
    does_not_include = "does not include",
    is = "is",
    is_not = "is not",
    is_empty = "is empty",
    is_not_empty = "is not empty",
    empty = "empty",
    the_answer = "the answer",
    this_answer = "this answer",
    number_selected_in = "number selected in",
    today = "today",
    now = "now",
    true = "true",
    false = "false",
    and = "and",
    or = "or",
    not = "not",
    repeat_for_each = "Repeat for each",
    instance_no = "Instance no.",
    other_specify = "Other (specify)",
    legend_required = "* Required question",
    legend_circle = "Choose one answer",
    legend_square = "Choose all answers that apply",
    legend_rank = paste("Rank the answers: write 1 in the box of the first",
                        "choice, 2 for the second, and so on"),
    legend_ask_if = "Ask if: ask the question only when the condition is met",
    not_on_paper = "Captured on the device, not on paper.",
    reference_no = "Reference no.",
    code = "Code",
    latitude = "Latitude",
    longitude = "Longitude",
    accuracy_m = "Accuracy (m)",
    range_steps = "%s to %s, steps of %s",
    range_from = "from",
    range_to = "to",
    at_least = "at least",
    at_most = "at most",
    more_than = "more than",
    less_than = "less than",
    fmt_date = "dd/mm/yyyy",
    fmt_month_year = "mm/yyyy",
    fmt_year = "yyyy",
    fmt_time = "hh:mm",
    choices_from = "Choices from",
    calculated = "Calculated",
    page = "Page",
    of = "of",
    version = "Version",
    generated = "Generated",
    form_id = "Form ID",
    language = "Language",
    image = "image",
    go_to = "Go to",
    go_to_end = "Go to end",
    legend_go_to = "Go to: skip to the question shown when that answer is given",
    entity_list = "Entity list"
  )
}

#' Built-in string table for French
#' @noRd
.strings_fr <- function() {
  list(
    ask_if = "Poser si",
    constraint = "Contrainte",
    choice_filter = "Choix propos\u00e9s si",
    includes = "comprend",
    does_not_include = "ne comprend pas",
    is = "est",
    is_not = "n'est pas",
    is_empty = "est vide",
    is_not_empty = "n'est pas vide",
    empty = "vide",
    the_answer = "la r\u00e9ponse",
    this_answer = "cette r\u00e9ponse",
    number_selected_in = "nombre de choix s\u00e9lectionn\u00e9s dans",
    today = "aujourd'hui",
    now = "maintenant",
    true = "vrai",
    false = "faux",
    and = "et",
    or = "ou",
    not = "non",
    repeat_for_each = "R\u00e9p\u00e9ter pour chaque",
    instance_no = "Occurrence n\u00b0",
    other_specify = "Autre (pr\u00e9ciser)",
    legend_required = "* Question obligatoire",
    legend_circle = "Choisir une seule r\u00e9ponse",
    legend_square = "Cocher toutes les r\u00e9ponses applicables",
    legend_rank = paste("Classer les r\u00e9ponses : \u00e9crire 1 dans la",
                        "case du premier choix, 2 pour le deuxi\u00e8me,",
                        "et ainsi de suite"),
    legend_ask_if = paste("Poser si : ne poser la question que si la",
                          "condition est remplie"),
    not_on_paper = "Saisi sur l'appareil, pas sur papier.",
    reference_no = "R\u00e9f\u00e9rence n\u00b0",
    code = "Code",
    latitude = "Latitude",
    longitude = "Longitude",
    accuracy_m = "Pr\u00e9cision (m)",
    range_steps = "de %s \u00e0 %s, par pas de %s",
    range_from = "de",
    range_to = "\u00e0",
    at_least = "au moins",
    at_most = "au plus",
    more_than = "plus de",
    less_than = "moins de",
    fmt_date = "jj/mm/aaaa",
    fmt_month_year = "mm/aaaa",
    fmt_year = "aaaa",
    fmt_time = "hh:mm",
    choices_from = "Choix provenant de",
    calculated = "Calcul\u00e9",
    page = "Page",
    of = "sur",
    version = "Version",
    generated = "G\u00e9n\u00e9r\u00e9 le",
    form_id = "Identifiant du formulaire",
    language = "Langue",
    image = "image",
    go_to = "Aller \u00e0",
    go_to_end = "Aller \u00e0 la fin",
    legend_go_to = paste("Aller \u00e0 : passer \u00e0 la question indiqu\u00e9e",
                         "quand cette r\u00e9ponse est donn\u00e9e"),
    entity_list = "Liste d'entit\u00e9s"
  )
}

#' Built-in string table for Spanish (machine-drafted, needs review)
#' @noRd
.strings_es <- function() {
  list(
    ask_if = "Preguntar si",
    constraint = "Restricci\u00f3n",
    choice_filter = "Opciones mostradas si",
    includes = "incluye",
    does_not_include = "no incluye",
    is = "es",
    is_not = "no es",
    is_empty = "est\u00e1 en blanco",
    is_not_empty = "no est\u00e1 en blanco",
    empty = "en blanco",
    the_answer = "la respuesta",
    this_answer = "esta respuesta",
    number_selected_in = "n\u00famero de opciones marcadas en",
    today = "hoy",
    now = "ahora",
    true = "verdadero",
    false = "falso",
    and = "y",
    or = "o",
    not = "no",
    repeat_for_each = "Repetir para cada",
    instance_no = "Repetici\u00f3n n.\u00ba",
    other_specify = "Otro (especificar)",
    legend_required = "* Pregunta obligatoria",
    legend_circle = "Elija una sola respuesta",
    legend_square = "Marque todas las respuestas que correspondan",
    legend_rank = paste(
      "Ordene las respuestas: escriba 1 en la casilla de la primera",
      "opci\u00f3n, 2 en la de la segunda, y as\u00ed sucesivamente"),
    legend_ask_if =
      "Preguntar si: hacer la pregunta solo cuando se cumple la condici\u00f3n",
    not_on_paper = "Se registra en el dispositivo, no en papel.",
    reference_no = "N.\u00ba de referencia",
    code = "C\u00f3digo",
    latitude = "Latitud",
    longitude = "Longitud",
    accuracy_m = "Precisi\u00f3n (m)",
    range_steps = "de %s a %s, en pasos de %s",
    range_from = "de",
    range_to = "a",
    at_least = "al menos",
    at_most = "como m\u00e1ximo",
    more_than = "m\u00e1s de",
    less_than = "menos de",
    fmt_date = "dd/mm/aaaa",
    fmt_month_year = "mm/aaaa",
    fmt_year = "aaaa",
    fmt_time = "hh:mm",
    choices_from = "Opciones de",
    calculated = "Calculado",
    page = "P\u00e1gina",
    of = "de",
    version = "Versi\u00f3n",
    generated = "Generado",
    form_id = "ID del formulario",
    language = "Idioma",
    image = "imagen",
    go_to = "Ir a",
    go_to_end = "Ir al final",
    legend_go_to =
      "Ir a: pasar a la pregunta indicada cuando se da esa respuesta",
    entity_list = "Lista de entidades"
  )
}

#' Built-in string table for Portuguese (machine-drafted, needs review)
#' @noRd
.strings_pt <- function() {
  list(
    ask_if = "Perguntar se",
    constraint = "Restri\u00e7\u00e3o",
    choice_filter = "Op\u00e7\u00f5es mostradas se",
    includes = "inclui",
    does_not_include = "n\u00e3o inclui",
    is = "\u00e9",
    is_not = "n\u00e3o \u00e9",
    is_empty = "est\u00e1 em branco",
    is_not_empty = "n\u00e3o est\u00e1 em branco",
    empty = "em branco",
    the_answer = "a resposta",
    this_answer = "esta resposta",
    number_selected_in = "n\u00famero de op\u00e7\u00f5es marcadas em",
    today = "hoje",
    now = "agora",
    true = "verdadeiro",
    false = "falso",
    and = "e",
    or = "ou",
    not = "n\u00e3o",
    repeat_for_each = "Repetir para cada",
    instance_no = "Repeti\u00e7\u00e3o n.\u00ba",
    other_specify = "Outro (especificar)",
    legend_required = "* Pergunta obrigat\u00f3ria",
    legend_circle = "Escolha uma s\u00f3 resposta",
    legend_square = "Marque todas as respostas que se aplicam",
    legend_rank = paste(
      "Ordene as respostas: escreva 1 na caixa da primeira op\u00e7\u00e3o, 2",
      "na da segunda, e assim por diante"),
    legend_ask_if = paste(
      "Perguntar se: fazer a pergunta s\u00f3 quando a condi\u00e7\u00e3o",
      "\u00e9 cumprida"),
    not_on_paper = "Preenchido no dispositivo, n\u00e3o em papel.",
    reference_no = "N.\u00ba de refer\u00eancia",
    code = "C\u00f3digo",
    latitude = "Latitude",
    longitude = "Longitude",
    accuracy_m = "Precis\u00e3o (m)",
    range_steps = "de %s a %s, em passos de %s",
    range_from = "de",
    range_to = "a",
    at_least = "pelo menos",
    at_most = "no m\u00e1ximo",
    more_than = "mais de",
    less_than = "menos de",
    fmt_date = "dd/mm/aaaa",
    fmt_month_year = "mm/aaaa",
    fmt_year = "aaaa",
    fmt_time = "hh:mm",
    choices_from = "Op\u00e7\u00f5es de",
    calculated = "Calculado",
    page = "P\u00e1gina",
    of = "de",
    version = "Vers\u00e3o",
    generated = "Gerado",
    form_id = "ID do formul\u00e1rio",
    language = "Idioma",
    image = "imagem",
    go_to = "Ir para",
    go_to_end = "Ir para o fim",
    legend_go_to = paste(
      "Ir para: passar \u00e0 pergunta indicada quando essa resposta \u00e9",
      "dada"),
    entity_list = "Lista de entidades"
  )
}

#' Built-in string table for Arabic (machine-drafted, needs review)
#' @noRd
.strings_ar <- function() {
  list(
    ask_if = "\u0627\u0633\u0623\u0644 \u0625\u0630\u0627",
    constraint = "\u0642\u064a\u062f",
    choice_filter = paste(
      "\u062a\u0638\u0647\u0631",
      "\u0627\u0644\u062e\u064a\u0627\u0631\u0627\u062a \u0625\u0630\u0627"),
    includes = "\u064a\u062a\u0636\u0645\u0646",
    does_not_include = "\u0644\u0627 \u064a\u062a\u0636\u0645\u0646",
    is = "\u064a\u0633\u0627\u0648\u064a",
    is_not = "\u0644\u0627 \u064a\u0633\u0627\u0648\u064a",
    is_empty = "\u0641\u0627\u0631\u063a",
    is_not_empty = "\u063a\u064a\u0631 \u0641\u0627\u0631\u063a",
    empty = "\u0641\u0627\u0631\u063a",
    the_answer = "\u0627\u0644\u0625\u062c\u0627\u0628\u0629",
    this_answer =
      "\u0647\u0630\u0647 \u0627\u0644\u0625\u062c\u0627\u0628\u0629",
    number_selected_in = paste(
      "\u0639\u062f\u062f \u0627\u0644\u062e\u064a\u0627\u0631\u0627\u062a",
      "\u0627\u0644\u0645\u062d\u062f\u062f\u0629 \u0641\u064a"),
    today = "\u0627\u0644\u064a\u0648\u0645",
    now = "\u0627\u0644\u0622\u0646",
    true = "\u0635\u062d\u064a\u062d",
    false = "\u062e\u0637\u0623",
    and = "\u0648",
    or = "\u0623\u0648",
    not = "\u0644\u064a\u0633",
    repeat_for_each = "\u0643\u0631\u0631 \u0644\u0643\u0644",
    instance_no =
      "\u0631\u0642\u0645 \u0627\u0644\u062a\u0643\u0631\u0627\u0631",
    other_specify = "\u0623\u062e\u0631\u0649 (\u062d\u062f\u062f)",
    legend_required =
      "* \u0633\u0624\u0627\u0644 \u0625\u0644\u0632\u0627\u0645\u064a",
    legend_circle = paste(
      "\u0627\u062e\u062a\u0631 \u0625\u062c\u0627\u0628\u0629",
      "\u0648\u0627\u062d\u062f\u0629"),
    legend_square = paste(
      "\u0627\u062e\u062a\u0631 \u0643\u0644",
      "\u0627\u0644\u0625\u062c\u0627\u0628\u0627\u062a",
      "\u0627\u0644\u0645\u0646\u0637\u0628\u0642\u0629"),
    legend_rank = paste(
      "\u0631\u062a\u0628 \u0627\u0644\u0625\u062c\u0627\u0628\u0627\u062a:",
      "\u0627\u0643\u062a\u0628 1 \u0641\u064a \u0645\u0631\u0628\u0639",
      "\u0627\u0644\u062e\u064a\u0627\u0631",
      "\u0627\u0644\u0623\u0648\u0644\u060c \u06482",
      "\u0644\u0644\u062e\u064a\u0627\u0631",
      "\u0627\u0644\u062b\u0627\u0646\u064a\u060c",
      "\u0648\u0647\u0643\u0630\u0627"),
    legend_ask_if = paste(
      "\u0627\u0633\u0623\u0644 \u0625\u0630\u0627: \u0627\u0637\u0631\u062d",
      "\u0627\u0644\u0633\u0624\u0627\u0644 \u0641\u0642\u0637",
      "\u0639\u0646\u062f \u062a\u062d\u0642\u0642",
      "\u0627\u0644\u0634\u0631\u0637"),
    not_on_paper = paste(
      "\u064a\u064f\u0633\u062c\u064e\u0651\u0644 \u0639\u0644\u0649",
      "\u0627\u0644\u062c\u0647\u0627\u0632\u060c \u0648\u0644\u064a\u0633",
      "\u0639\u0644\u0649 \u0627\u0644\u0648\u0631\u0642."),
    reference_no = paste(
      "\u0627\u0644\u0631\u0642\u0645",
      "\u0627\u0644\u0645\u0631\u062c\u0639\u064a"),
    code = "\u0627\u0644\u0631\u0645\u0632",
    latitude = "\u062e\u0637 \u0627\u0644\u0639\u0631\u0636",
    longitude = "\u062e\u0637 \u0627\u0644\u0637\u0648\u0644",
    accuracy_m = "\u0627\u0644\u062f\u0642\u0629 (\u0645)",
    range_steps = paste(
      "\u0645\u0646 %s \u0625\u0644\u0649 %s\u060c",
      "\u0628\u062e\u0637\u0648\u0627\u062a \u0642\u062f\u0631\u0647\u0627 %s"),
    range_from = "\u0645\u0646",
    range_to = "\u0625\u0644\u0649",
    at_least = "\u0644\u0627 \u064a\u0642\u0644 \u0639\u0646",
    at_most = "\u0644\u0627 \u064a\u0632\u064a\u062f \u0639\u0646",
    more_than = "\u0623\u0643\u062b\u0631 \u0645\u0646",
    less_than = "\u0623\u0642\u0644 \u0645\u0646",
    fmt_date = "dd/mm/yyyy",
    fmt_month_year = "mm/yyyy",
    fmt_year = "yyyy",
    fmt_time = "hh:mm",
    choices_from =
      "\u0627\u0644\u062e\u064a\u0627\u0631\u0627\u062a \u0645\u0646",
    calculated = "\u0645\u062d\u0633\u0648\u0628",
    page = "\u0635\u0641\u062d\u0629",
    of = "\u0645\u0646",
    version = "\u0627\u0644\u0625\u0635\u062f\u0627\u0631",
    generated = paste(
      "\u062a\u0627\u0631\u064a\u062e",
      "\u0627\u0644\u0625\u0646\u0634\u0627\u0621"),
    form_id = paste(
      "\u0645\u0639\u0631\u0651\u0641",
      "\u0627\u0644\u0627\u0633\u062a\u0645\u0627\u0631\u0629"),
    language = "\u0627\u0644\u0644\u063a\u0629",
    image = "\u0635\u0648\u0631\u0629",
    go_to = "\u0627\u0646\u062a\u0642\u0644 \u0625\u0644\u0649",
    go_to_end = paste(
      "\u0627\u0646\u062a\u0642\u0644 \u0625\u0644\u0649",
      "\u0627\u0644\u0646\u0647\u0627\u064a\u0629"),
    legend_go_to = paste(
      "\u0627\u0646\u062a\u0642\u0644 \u0625\u0644\u0649:",
      "\u0627\u0646\u062a\u0642\u0644 \u0625\u0644\u0649",
      "\u0627\u0644\u0633\u0624\u0627\u0644",
      "\u0627\u0644\u0645\u0628\u064a\u0646 \u0639\u0646\u062f",
      "\u0625\u0639\u0637\u0627\u0621 \u062a\u0644\u0643",
      "\u0627\u0644\u0625\u062c\u0627\u0628\u0629"),
    entity_list = paste(
      "\u0642\u0627\u0626\u0645\u0629",
      "\u0627\u0644\u0643\u064a\u0627\u0646\u0627\u062a")
  )
}

#' Built-in string table for Bengali (machine-drafted, needs review)
#' @noRd
.strings_bn <- function() {
  list(
    ask_if = paste(
      "\u099c\u09bf\u099c\u09cd\u099e\u09be\u09b8\u09be",
      "\u0995\u09b0\u09c1\u09a8 \u09af\u09a6\u09bf"),
    constraint = "\u09b8\u09c0\u09ae\u09be\u09ac\u09a6\u09cd\u09a7\u09a4\u09be",
    choice_filter = paste(
      "\u09ac\u09bf\u0995\u09b2\u09cd\u09aa \u09a6\u09c7\u0996\u09be\u09a8",
      "\u09af\u09a6\u09bf"),
    includes = "\u098f\u09b0 \u09ae\u09a7\u09cd\u09af\u09c7 \u0986\u099b\u09c7",
    does_not_include =
      "\u098f\u09b0 \u09ae\u09a7\u09cd\u09af\u09c7 \u09a8\u09c7\u0987",
    is = "\u09b9\u09b2\u09cb",
    is_not = "\u09a8\u09af\u09bc",
    is_empty = "\u09ab\u09be\u0981\u0995\u09be",
    is_not_empty = "\u09ab\u09be\u0981\u0995\u09be \u09a8\u09af\u09bc",
    empty = "\u09ab\u09be\u0981\u0995\u09be",
    the_answer = "\u0989\u09a4\u09cd\u09a4\u09b0",
    this_answer = "\u098f\u0987 \u0989\u09a4\u09cd\u09a4\u09b0",
    number_selected_in = paste(
      "\u09a8\u09bf\u09b0\u09cd\u09ac\u09be\u099a\u09bf\u09a4",
      "\u09ac\u09bf\u0995\u09b2\u09cd\u09aa\u09c7\u09b0",
      "\u09b8\u0982\u0996\u09cd\u09af\u09be"),
    today = "\u0986\u099c",
    now = "\u098f\u0996\u09a8",
    true = "\u09b8\u09a4\u09cd\u09af",
    false = "\u09ae\u09bf\u09a5\u09cd\u09af\u09be",
    and = "\u098f\u09ac\u0982",
    or = "\u0985\u09a5\u09ac\u09be",
    not = "\u09a8\u09af\u09bc",
    repeat_for_each = paste(
      "\u09aa\u09cd\u09b0\u09a4\u09bf\u099f\u09bf\u09b0",
      "\u099c\u09a8\u09cd\u09af",
      "\u09aa\u09c1\u09a8\u09b0\u09be\u09ac\u09c3\u09a4\u09cd\u09a4\u09bf",
      "\u0995\u09b0\u09c1\u09a8"),
    instance_no = "\u0995\u09cd\u09b0\u09ae\u09bf\u0995 \u09a8\u0982",
    other_specify = paste(
      "\u0985\u09a8\u09cd\u09af\u09be\u09a8\u09cd\u09af",
      "(\u0989\u09b2\u09cd\u09b2\u09c7\u0996 \u0995\u09b0\u09c1\u09a8)"),
    legend_required = paste(
      "* \u0986\u09ac\u09b6\u09cd\u09af\u09bf\u0995",
      "\u09aa\u09cd\u09b0\u09b6\u09cd\u09a8"),
    legend_circle = paste(
      "\u098f\u0995\u099f\u09bf \u0989\u09a4\u09cd\u09a4\u09b0",
      "\u09ac\u09c7\u099b\u09c7 \u09a8\u09bf\u09a8"),
    legend_square = paste(
      "\u09aa\u09cd\u09b0\u09af\u09cb\u099c\u09cd\u09af \u09b8\u09ac",
      "\u0989\u09a4\u09cd\u09a4\u09b0 \u09ac\u09c7\u099b\u09c7",
      "\u09a8\u09bf\u09a8"),
    legend_rank = paste(
      "\u0989\u09a4\u09cd\u09a4\u09b0\u0997\u09c1\u09b2\u09cb\u09b0",
      "\u0995\u09cd\u09b0\u09ae \u09a6\u09bf\u09a8:",
      "\u09aa\u09cd\u09b0\u09a5\u09ae",
      "\u09aa\u099b\u09a8\u09cd\u09a6\u09c7\u09b0 \u0998\u09b0\u09c7 1",
      "\u09b2\u09bf\u0996\u09c1\u09a8,",
      "\u09a6\u09cd\u09ac\u09bf\u09a4\u09c0\u09af\u09bc\u099f\u09bf\u09b0",
      "\u0998\u09b0\u09c7 2, \u098f\u09ad\u09be\u09ac\u09c7",
      "\u09ac\u09be\u0995\u09bf\u0997\u09c1\u09b2\u09cb"),
    legend_ask_if = paste(
      "\u099c\u09bf\u099c\u09cd\u099e\u09be\u09b8\u09be",
      "\u0995\u09b0\u09c1\u09a8 \u09af\u09a6\u09bf: \u09b6\u09b0\u09cd\u09a4",
      "\u09aa\u09c2\u09b0\u09a3 \u09b9\u09b2\u09c7\u0987",
      "\u0995\u09c7\u09ac\u09b2",
      "\u09aa\u09cd\u09b0\u09b6\u09cd\u09a8\u099f\u09bf",
      "\u099c\u09bf\u099c\u09cd\u099e\u09be\u09b8\u09be",
      "\u0995\u09b0\u09c1\u09a8"),
    not_on_paper = paste(
      "\u09a1\u09bf\u09ad\u09be\u0987\u09b8\u09c7",
      "\u09b8\u0982\u0997\u09cd\u09b0\u09b9 \u0995\u09b0\u09be",
      "\u09b9\u09af\u09bc, \u0995\u09be\u0997\u099c\u09c7",
      "\u09a8\u09af\u09bc\u0964"),
    reference_no =
      "\u09b0\u09c7\u09ab\u09be\u09b0\u09c7\u09a8\u09cd\u09b8 \u09a8\u0982",
    code = "\u0995\u09cb\u09a1",
    latitude = "\u0985\u0995\u09cd\u09b7\u09be\u0982\u09b6",
    longitude = "\u09a6\u09cd\u09b0\u09be\u0998\u09bf\u09ae\u09be\u0982\u09b6",
    accuracy_m =
      "\u09a8\u09bf\u09b0\u09cd\u09ad\u09c1\u09b2\u09a4\u09be (\u09ae\u09bf.)",
    range_steps = "%s \u09a5\u09c7\u0995\u09c7 %s, %s \u0995\u09b0\u09c7",
    range_from = "\u09a8\u09cd\u09af\u09c2\u09a8\u09a4\u09ae",
    range_to = "\u0993 \u09b8\u09b0\u09cd\u09ac\u09cb\u099a\u09cd\u099a",
    at_least = "\u09a8\u09cd\u09af\u09c2\u09a8\u09a4\u09ae",
    at_most = "\u09b8\u09b0\u09cd\u09ac\u09cb\u099a\u09cd\u099a",
    more_than = ">",
    less_than = "<",
    fmt_date = "dd/mm/yyyy",
    fmt_month_year = "mm/yyyy",
    fmt_year = "yyyy",
    fmt_time = "hh:mm",
    choices_from =
      "\u09ac\u09bf\u0995\u09b2\u09cd\u09aa\u09c7\u09b0 \u0989\u09ce\u09b8",
    calculated = "\u0997\u09a3\u09a8\u09be\u0995\u09c3\u09a4",
    page = "\u09aa\u09c3\u09b7\u09cd\u09a0\u09be",
    of = "/",
    version = "\u09b8\u0982\u09b8\u09cd\u0995\u09b0\u09a3",
    generated = "\u09a4\u09c8\u09b0\u09bf",
    form_id = "\u09ab\u09b0\u09cd\u09ae \u0986\u0987\u09a1\u09bf",
    language = "\u09ad\u09be\u09b7\u09be",
    image = "\u099b\u09ac\u09bf",
    go_to = "\u09af\u09be\u09a8",
    go_to_end = "\u09b6\u09c7\u09b7\u09c7 \u09af\u09be\u09a8",
    legend_go_to = paste(
      "\u09af\u09be\u09a8: \u0989\u09a4\u09cd\u09a4\u09b0\u099f\u09bf",
      "\u09a6\u09c7\u0993\u09af\u09bc\u09be \u09b9\u09b2\u09c7",
      "\u09a6\u09c7\u0996\u09be\u09a8\u09cb",
      "\u09aa\u09cd\u09b0\u09b6\u09cd\u09a8\u09c7 \u099a\u09b2\u09c7",
      "\u09af\u09be\u09a8"),
    entity_list = paste(
      "\u098f\u09a8\u099f\u09bf\u099f\u09bf",
      "\u09a4\u09be\u09b2\u09bf\u0995\u09be")
  )
}

#' Built-in string table for Swahili (machine-drafted, needs review)
#' @noRd
.strings_sw <- function() {
  list(
    ask_if = "Uliza kama",
    constraint = "Sharti",
    choice_filter = "Machaguo yanaonyeshwa kama",
    includes = "inajumuisha",
    does_not_include = "haijumuishi",
    is = "ni",
    is_not = "si",
    is_empty = "ni tupu",
    is_not_empty = "si tupu",
    empty = "tupu",
    the_answer = "jibu",
    this_answer = "jibu hili",
    number_selected_in = "idadi ya machaguo yaliyochaguliwa katika",
    today = "leo",
    now = "sasa",
    true = "kweli",
    false = "si kweli",
    and = "na",
    or = "au",
    not = "si",
    repeat_for_each = "Rudia kwa kila",
    instance_no = "Marudio Na.",
    other_specify = "Nyingine (taja)",
    legend_required = "* Swali la lazima",
    legend_circle = "Chagua jibu moja",
    legend_square = "Chagua majibu yote yanayohusika",
    legend_rank = paste(
      "Panga majibu: andika 1 kwenye kisanduku cha chaguo la kwanza, 2 kwa la",
      "pili, na kuendelea"),
    legend_ask_if = "Uliza kama: uliza swali tu pale sharti linapotimizwa",
    not_on_paper = "Hujazwa kwenye kifaa, si kwenye karatasi.",
    reference_no = "Nambari ya kumbukumbu",
    code = "Msimbo",
    latitude = "Latitudo",
    longitude = "Longitudo",
    accuracy_m = "Usahihi (m)",
    range_steps = "%s hadi %s, kwa hatua za %s",
    range_from = "kuanzia",
    range_to = "hadi",
    at_least = "angalau",
    at_most = "isiyozidi",
    more_than = "zaidi ya",
    less_than = "chini ya",
    fmt_date = "dd/mm/yyyy",
    fmt_month_year = "mm/yyyy",
    fmt_year = "yyyy",
    fmt_time = "hh:mm",
    choices_from = "Machaguo kutoka",
    calculated = "Hukokotolewa",
    page = "Ukurasa",
    of = "kati ya",
    version = "Toleo",
    generated = "Imetengenezwa",
    form_id = "Kitambulisho cha fomu",
    language = "Lugha",
    image = "picha",
    go_to = "Nenda kwa",
    go_to_end = "Nenda mwisho",
    legend_go_to =
      "Nenda kwa: ruka hadi swali lililoonyeshwa jibu hilo likitolewa",
    entity_list = "Orodha ya huluki"
  )
}

#' Names of the strings that can be overridden
#' @noRd
.strings_names <- function() {
  names(.strings_en())
}

#' The string table for a language
#'
#' @param lang_code Language code such as `"fr"` (NA, NULL or an unknown code
#'   gives English). Only the primary subtag is used, so `"fr-CA"` is French.
#' @param overrides NULL or a named list of single strings that replace
#'   entries of the table. Unknown names are an error.
#' @return A named list of strings.
#' @noRd
.cxf_strings <- function(lang_code = NULL, overrides = NULL) {
  code <- if (is.null(lang_code) || length(lang_code) != 1L ||
              is.na(lang_code)) {
    "en"
  } else {
    sub("[-_].*$", "", tolower(trimws(lang_code)))
  }
  out <- switch(code, fr = .strings_fr(), es = .strings_es(),
                pt = .strings_pt(), ar = .strings_ar(), bn = .strings_bn(),
                sw = .strings_sw(), .strings_en())
  if (!is.null(overrides)) {
    bad <- setdiff(names(overrides), names(out))
    if (length(bad) > 0L || is.null(names(overrides))) {
      valid_txt <- paste(names(out), collapse = ", ")
      .cxf_abort(c(
        "Unknown string name{?s} in {.arg strings}: {.val {bad}}.",
        i = "Valid names are: {valid_txt}."
      ))
    }
    for (nm in names(overrides)) {
      out[[nm]] <- overrides[[nm]]
    }
  }
  out
}

#' Language code of the form in use ("en" when unknown)
#'
#' Taken from the `code` column of `form$languages` for the chosen suffix. A
#' name alone (for example "French" or "Kiswahili") is looked up in
#' `.lang_name_codes()`.
#' @noRd
.form_lang_code <- function(form) {
  sfx <- form$lang$suffix
  langs <- form$languages
  if (is.null(sfx) || is.na(sfx) || is.null(langs) || nrow(langs) == 0L) {
    return("en")
  }
  i <- match(tolower(sfx), tolower(langs$suffix))
  if (is.na(i)) {
    return("en")
  }
  code <- langs$code[i]
  if (is.na(code)) {
    nm <- langs$name[i]
    if (is.na(nm)) {
      return("en")
    }
    code <- unname(.lang_name_codes()[.fold_name(nm)])
    if (is.na(code)) {
      return("en")
    }
  }
  sub("[-_].*$", "", tolower(code))
}
