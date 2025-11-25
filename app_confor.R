# app.R - Partie 1/3 (Imports, Styles, Utilitaires, Render Question)

# --- Packages requis ---
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr); library(tidyr)
  library(purrr);  library(officer); library(flextable); library(glue); library(tools)
  library(tibble); library(rlang); library(shiny); library(DT); library(writexl); library(htmltools)
})

# =====================================================================
# XLSForm -> Word (Rendu Word) - Script corrigé V2
# =====================================================================

# Opérateur de coalescence robuste corrigé pour gérer les vecteurs
`%||%` <- function(a, b) {
  if (is.null(a) || length(a) == 0 || (is.atomic(a) && length(a) == 1 && (is.na(a) || !nzchar(a)))) {
    b
  } else {
    a
  }
}

# ---------------------------------------------------------------------
# PARAMÈTRES DE STYLE POUR UN RENDU PROFESSIONNEL
# ---------------------------------------------------------------------
WFP_BLUE    <- "#0A66C2"   
GREY_TXT    <- "#777777"
WFP_DARK_BLUE <- "#001F3F" 
GREY_BG     <- "#F2F2F2"   
WHITE_TXT   <- "#FFFFFF"   
RED_TXT     <- "#C00000" 
FONT_FAMILY <- "Cambria (Body)"     
LINE_SP     <- 1.0        
INDENT_Q    <- 0.3        
INDENT_C    <- 0.5        

FS_TITLE    <- 14
FS_SUBTITLE <- 12
FS_BLOCK    <- 12
FS_META     <- 10    
FS_HINT     <- 9   
FS_RELV     <- 9   
FS_MISC     <- 10     
FS_Q_BLUE   <- 11  
FS_Q_GREY   <- 9   

EXCLUDE_TYPES <- c("calculate","start","end","today","deviceid","subscriberid","simserial","phonenumber","username","instanceid","end_group", "end group")
EXCLUDE_NAMES <- c("start","end","today","deviceid","username","instanceID","instanceid")

LOGO_PATH  <- "cp_logo.png"

# ---------------------------------------------------------------------
# Fonctions utilitaires 
# ---------------------------------------------------------------------

generate_output_path <- function(base_path) {
  if (!file.exists(base_path)) return(base_path)
  path_dir <- dirname(base_path)
  path_ext <- tools::file_ext(base_path)
  path_name <- tools::file_path_sans_ext(base_path)
  i <- 1
  new_path <- base_path
  while (file.exists(new_path)) {
    new_name <- paste0(path_name, "_", i, ".", path_ext)
    new_path <- file.path(path_dir, basename(new_name)) 
    i <- i + 1
  }
  return(new_path)
}

# Fonctions de détection de colonne mises à jour pour la langue et robustesse
detect_label_col <- function(df, target_language = NULL) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  if (!is.null(target_language) && target_language != "Auto") {
    specific_label <- cols[grepl(glue("label(::|:){target_language}"), cols_clean, ignore.case = TRUE)]
    if (length(specific_label) > 0) return(specific_label)
  }
  
  fr <- cols[grepl("label(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  ll <- cols[grepl("label(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  if ("label" %in% cols_clean) return(cols[cols_clean == "label"])
  NA_character_
}

detect_hint_col <- function(df, target_language = NULL) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  if (!is.null(target_language) && target_language != "Auto") {
    specific_hint <- cols[grepl(glue("hint(::|:){target_language}"), cols_clean, ignore.case = TRUE)]
    if (length(specific_hint) > 0) return(specific_hint)
  }
  
  fr <- cols[grepl("hint(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  if ("hint" %in% cols_clean) return(cols[cols_clean == "hint"])
  ll <- cols[grepl("hint(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  if ("hint" %in% cols_clean) return(cols[cols_clean == "hint"])
  NA_character_
}

# Fonction translate_relevant complète
translate_relevant <- function(expr, labels, choices, df_survey) {
  if (is.null(expr) || is.na(expr) || !nzchar(trimws(expr))) return(NA_character_)
  txt <- expr
  
  choices <- choices %>% mutate(
    name = tolower(str_trim(as.character(name))),
    list_name = tolower(str_trim(as.character(list_name)))
  )
  df_survey_clean <- df_survey %>% mutate(
    name = tolower(str_trim(as.character(name))),
    type = tolower(as.character(type))
  )
  
  get_choice_label <- function(code, list_name){
    code_clean <- tolower(str_trim(as.character(code)))
    list_name_clean <- tolower(str_trim(as.character(list_name)))
    r <- choices %>% filter(name == code_clean & list_name == list_name_clean)
    if (nrow(r) > 0) {
      label_col_name_choices <- detect_label_col(choices)
      if (!is.na(label_col_name_choices)) {
        lab_val <- as.character(r[[label_col_name_choices]]) 
        lab <- lab_val %||% as.character(code)
      } else {
        lab <- as.character(r$name) %||% as.character(code)
      }
      return(str_replace_all(lab, "<.*?>", ""))
    } else { return(as.character(code)) }
  }
  
  get_var_label <- function(v){ 
    val <- tryCatch(labels[[v]], error = function(e) NULL)
    if(is.null(val) || is.na(val) || !nzchar(val)) {
      clean_val <- as.character(v)
    } else {
      clean_val <- str_replace_all(as.character(val), "<.*?>", "")
    }
    return(clean_val) 
  }
  
  get_listname_from_varname <- function(var_name, survey_df) {
    var_name_clean <- tolower(str_trim(as.character(var_name)))
    row <- survey_df %>% filter(name == var_name_clean) %>% head(1)
    if (nrow(row) > 0) {
      type_val <- tolower(as.character(row$type))
      list_name <- stringr::str_replace(type_val, "^select_(one|multiple)\\s+", "")
      final_list_name <- ifelse(
        test = list_name == type_val,
        yes = as.character(row$list_name %||% NA_character_),
        no = list_name
      )
      return(str_trim(final_list_name))
    } else { 
      return(NA_character_) 
    }
  }
  
  txt <- stringr::str_replace_all(txt, "selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+'\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey_clean)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("`%s` a l'option «%s» cochée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  txt <- stringr::str_replace_all(txt, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+?'\\s*\\)\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey_clean)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("`%s` N'A PAS l'option «%s» cochée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  comparison_regex <- "\\$\\{[^}]+\\}\\s*[=><!]+\\s*('|\")?[^'\"]+('|\")?"
  txt <- stringr::str_replace_all(txt, comparison_regex, function(comp_expr) {
    var_match <- str_match(comp_expr, "\\$\\{([^}]+)\\}")
    var_name <- var_match[, 2]
    code_match <- str_match(comp_expr, "('|\")?([^'\"]+)('|\")?$")
    code_val <- code_match[, 3]
    current_listname <- get_listname_from_varname(var_name, df_survey_clean)
    choice_label <- get_choice_label(code_val, current_listname)
    reconstructed_expr <- sprintf("`%s` a la valeur «%s»", get_var_label(var_name), choice_label)
    reconstructed_expr <- str_replace(reconstructed_expr, "\\s*=\\s*", " est égal à ")
    reconstructed_expr <- str_replace(reconstructed_expr, "\\s*!=\\s*", " est différent de ")
    return(reconstructed_expr)
  })
  
  txt <- stringr::str_replace_all(txt, "\\$\\{[^}]+\\}", function(x){ 
    m <- stringr::str_match(x, "\\$\\{([^}]+)\\}")
    vars <- m[,2]
    vapply(vars, get_var_label, character(1)) 
  })
  
  txt <- str_replace_all(txt, "\\bandand\\b", " et "); 
  txt <- str_replace_all(txt, "\\bor\\b",  " ou "); 
  txt <- str_replace_all(txt, "\\bnot\\b", " non "); 
  txt <- str_replace_all(txt, "\\s*!=\\s*", " est différent de ");
  txt <- str_replace_all(txt, "\\s*=\\s*", " est égal à ");
  txt <- str_replace_all(txt, "\\s*>\\s*", " est supérieur à ");
  txt <- str_replace_all(txt, "\\s*>=\\s*", " est supérieur ou égal à ");
  txt <- str_replace_all(txt, "\\s*<\\s*", " est inférieur à ");
  txt <- str_replace_all(txt, "\\s*<=\\s*", " est inférieur ou égal à ");
  
  txt <- str_replace_all(txt, "count-selected\\s*\\(\\s*[^\\)]+\\)\\s*>=\\s*1", function(x){ 
    m <- str_match(x, "count-selected\\s*\\(\\s*(.+?)\\s*\\)\\s*>=\\s*1")
    v <- m[,2]
    vapply(v, function(z) sprintf("Au moins une option est cochée pour %s", z), character(1)) 
  })
  
  txt <- str_replace_all(txt, "\\s*\\(\\s*", "\\(") %>% str_replace_all("\\s*\\)\\s*", "\\)") %>% str_replace_all("\\s+", " ") %>% str_trim()
  return(txt)
}

# ---------------------------------------------------------------------
# Styles texte & paragraphe / Blocs visuels 
# ---------------------------------------------------------------------
fp_txt <- function(color="black", size=11, bold=FALSE, italic=FALSE, underline=FALSE) { fp_text(color = color, font.size = size, bold = bold, italic = italic, underline = underline, font.family = FONT_FAMILY) }
fp_q_blue   <- fp_txt(color = WFP_DARK_BLUE, size = FS_Q_BLUE, bold = FALSE) 
fp_q_grey   <- fp_txt(color = GREY_TXT, size = FS_Q_GREY)             
fp_sec_title <- fp_txt(color = WFP_BLUE, size = FS_TITLE, bold = TRUE, underline = TRUE)
fp_sub_title <- fp_txt(color = WFP_BLUE, size = FS_SUBTITLE, bold = TRUE, underline = TRUE)
fp_block     <- fp_txt(color = WFP_BLUE, size = FS_BLOCK,  bold = TRUE) 
fp_meta      <- fp_txt(color = GREY_TXT, size = FS_META)
fp_hint      <- fp_txt(color = GREY_TXT, size = FS_HINT, italic = TRUE)
fp_relevant  <- fp_txt(color = RED_TXT,  size = FS_RELV)
fp_missing_list <- fp_txt(color = GREY_TXT, size = FS_MISC, italic = TRUE)
p_default <- fp_par(text.align = "left", line_spacing = LINE_SP)  
p_q_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_Q * 72)
p_c_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_C * 72)

add_hrule <- function(doc, width = 1){ doc <- body_add_par(doc, ""); return(doc) }
add_band <- function(doc, text, txt_fp){ doc <- body_add_fpar(doc, fpar(ftext(text, txt_fp), fp_p = fp_par(padding.top = 4, padding.bottom = 4, padding.left = INDENT_Q*72))); return(doc) }

add_choice_lines <- function(doc, choices_map, list_name_to_filter, symbol = "○") {
  if (is.na(list_name_to_filter) || list_name_to_filter == "") return(doc)
  list_name_to_filter_clean <- str_replace_all(list_name_to_filter, "\\s", "")
  
  df <- choices_map[[list_name_to_filter_clean]]
  
  if (is.null(df) || nrow(df) == 0) {
    doc <- body_add_fpar(doc, fpar(ftext(glue("Commentaire: La liste de choix '{list_name_to_filter}' est introuvable ou vide dans l'onglet 'choices'."), fp_missing_list), fp_p = p_c_indent_fixed)) 
    return(doc)
  }
  
  df <- df %>% mutate(
    label_final = coalesce(label_col, as.character(name)),
    txt_final = sprintf("%s %s (%s)", symbol, label_final, name)
  )
  
  purrr::walk(df$txt_final, function(text_label) { 
    doc <<- body_add_fpar(doc, fpar(ftext(text_label, fp_txt(size = FS_MISC)), fp_p = p_c_indent_fixed)) 
  }); 
  return(doc)
}

add_placeholder_box <- function(doc, txt = "Réponse : [insérer votre réponse ici]"){
  doc <- body_add_fpar(doc, fpar(ftext(txt, fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_c_indent_fixed))
  doc <- body_add_par(doc, ""); return(doc)
}

# ---------------------------------------------------------------------
# Rendu d’une question
# ---------------------------------------------------------------------
render_question <- function(doc, row, number, label_col_name, hint_col_name, choices_map, lab_map, full_choices_sheet,full_survey_sheet){
  if (is.null(row) || nrow(row) == 0) return(doc)
  q_type <- tolower(row$type %||% "")
  q_name <- row$name %||% ""
  
  clean_html <- function(text) {
    text <- str_replace_all(text, "<[^>]+>", "")
    text <- str_replace_all(text, "&nbsp;", "")
    text <- str_replace_all(text, "\\s+", " ")
    return(str_trim(text))
  }
  
  q_lab  <- clean_html(row[[label_col_name]] %||% q_name)
  
  if (is.na(q_name) || q_name %in% EXCLUDE_NAMES) return(doc)
  
  if (q_type == "note") {
    ftext_blue <- ftext(sprintf("Note : %s", q_lab), fp_q_blue)
  } else {
    if (is.null(number) || is.na(number) || !nzchar(as.character(number))) return(doc)
    ftext_blue <- ftext(sprintf("%s. %s", number, q_lab), fp_q_blue)
  }
  ftext_grey <- ftext(sprintf(" (%s – %s)", q_name, q_type), fp_q_grey)
  doc <- body_add_fpar(doc, fpar(ftext_blue, ftext_grey, fp_p = p_q_indent_fixed)) 
  
  rel <- row$relevant %||% NA_character_
  h <- NA_character_
  if (!is.na(hint_col_name) && hint_col_name %in% names(row)) { 
    h <- row[[hint_col_name]] %||% NA_character_ 
    h <- clean_html(h)
  }
  
  if (!is.na(rel) && nzchar(rel)) { tr <- translate_relevant(rel, lab_map, full_choices_sheet, full_survey_sheet); doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) }
  if (!is.na(h) && nzchar(h)) { doc <- body_add_fpar(doc, fpar(ftext(h, fp_hint), fp_p = p_q_indent_fixed)) }
  
  if (str_starts(q_type, "select_one")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir la réponse parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_one\\s+", "", q_type));
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "○");
  } 
  else if (str_starts(q_type, "select_multiple")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir les réponses pertinentes parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_multiple\\s+", "", q_type)); 
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "☐"); 
  } 
  else if (str_detect(q_type, "^note")) { } 
  else { 
    placeholder <- "Réponse : [insérer votre réponse ici]"
    if (str_detect(q_type, "integer")) { placeholder <- "Réponse : [insérer un entier]" } 
    else if (str_detect(q_type, "decimal")) { placeholder <- "Réponse : [insérer un décimal]" }
    else if (str_detect(q_type, "date")) { placeholder <- "Réponse : [insérer une date]" }
    else if (str_detect(q_type, "geopoint")) { placeholder <- "Réponse : [capturer les coordonnées GPS]" }
    else if (str_detect(q_type, "image")) { placeholder <- "Réponse : [prendre une photo] ou inserer une image" }
    else if (str_detect(q_type, "photo")) { placeholder <- "Réponse : [prendre une photo]" }
    doc <- add_placeholder_box(doc, placeholder)
  }
  
  if (!str_detect(q_type, "^note|select_")) {
    doc <- body_add_fpar(doc, fpar(ftext("__________________________________________________________", fp_txt(size=FS_Q_BLUE)), fp_p = p_c_indent_fixed))
  }
  
  doc <- body_add_par(doc, "")
  return(doc)
}

# app.R - Partie 2/3 (Conversion Word adaptée et Validation d'erreurs)

# ---------------------------------------------------------------------
# Adaptation de xlsform_to_wordRev pour Shiny (CORRIGÉ ET COMPLET)
# ---------------------------------------------------------------------
xlsform_to_wordRev <- function(xlsx, output_path, selected_language = NULL) {
  
  # Lecture des données
  survey   <- read_excel(xlsx, sheet = "survey")
  choices  <- read_excel(xlsx, sheet = "choices")
  settings <- tryCatch(read_excel(xlsx, sheet = "settings"), error = function(e) NULL)
  
  # Détection de la colonne label/hint en fonction de la langue sélectionnée par l'utilisateur
  label_col_name <- detect_label_col(survey, target_language = selected_language)
  hint_col_name <- detect_hint_col(survey, target_language = selected_language)
  
  if (is.na(label_col_name)) stop("Colonne label introuvable pour la langue sélectionnée.")
  label_col_sym <- sym(label_col_name) 
  
  # Normalisation et préparation des données
  names(choices) <- tolower(iconv(names(choices), to = "ASCII//TRANSLIT"))
  doc_title <- if (!is.null(settings) && "form_title" %in% names(settings)) settings$form_title %||% "XLSForm – Rendu Word" else "XLSForm – Rendu Word"
  survey <- survey %>% filter(!duplicated(name))
  lab_map <- survey %>% select(name, !!label_col_sym) %>% mutate(name = as.character(name)) %>% tibble::deframe()
  
  # Initialisation du document Word (utilisation d'un document vierge pour Shiny)
  doc <- read_docx() 
  doc <- body_set_default_section(doc, prop_section(page_size = page_size(orient = "portrait", width = 8.5, height = 11), page_margins = page_mar(top = 1, bottom = 1, left = 1.0, right = 1.0, header = 0.5, footer = 0.5)))
  
  # Ajout du titre (le logo est optionnel et dépend du chemin local)
  fp_title_main <- fp_txt(color = RED_TXT, size = 16, bold = TRUE)
  p_title_main <- fp_par(text.align = "center", line_spacing = LINE_SP)
  
  # Assurez-vous que LOGO_PATH est défini dans la partie 1 si vous l'utilisez
  if (exists("LOGO_PATH") && !is.null(LOGO_PATH) && file.exists(LOGO_PATH)) { 
    doc <- body_add_par(doc, "", style = "Normal")
    doc <- body_add_fpar(doc, fpar(external_img(src = LOGO_PATH, height = 0.6, width = 0.6, unit = "in"), ftext("  "), ftext(doc_title, prop = fp_title_main), fp_p = p_title_main)) 
  } else { 
    doc <- body_add_fpar(doc, fpar(ftext(doc_title, fp_title_main), fp_p = p_title_main)) 
  }
  doc <- add_hrule(doc)
  
  # Préparation de la map des choix
  choices <- choices %>% 
    mutate(across(everything(), as.character)) %>% 
    mutate(across(everything(), ~ifelse(is.na(.), NA_character_, .))) %>%
    mutate(list_name = as.character(str_trim(str_replace_all(tolower(list_name), "[[:space:]]+", ""))))
  
  label_col_choices <- detect_label_col(choices, target_language = selected_language)
  if (is.na(label_col_choices)) stop("Colonne label choix introuvable pour la langue sélectionnée.")
  choices_map <- split(choices, choices$list_name)
  choices_map <- lapply(choices_map, function(df) { df %>% mutate(label_col = .data[[label_col_choices]]) })
  
  # Boucle principale (Logique de parcours du formulaire)
  sec_id <- 0L; sub_id <- 0L; q_id <- 0L; names_deja_traites <- character(0) 
  
  for (i in seq_len(nrow(survey))) {
    r <- survey[i, , drop = FALSE]
    t_raw <- as.character(r$type %||% "")
    t <- tolower(t_raw)
    qname <- as.character(r$name)
    if (!is.na(qname) && nzchar(qname)) { if (qname %in% names_deja_traites) { next } else { names_deja_traites <- c(names_deja_traites, qname) } }
    
    if (t %in% EXCLUDE_TYPES) next
    if (!is.null(r$name) && any(tolower(r$name) %in% tolower(EXCLUDE_NAMES))) next
    
    current_number <- NULL
    is_group <- str_starts(t, "begin_group") && !str_starts(t, "begin_repeat")
    is_repeat <- str_starts(t, "begin_repeat") || str_starts(t, "begin repeat")
    is_repeat_end <- str_starts(t, "end_repeat") || str_starts(t, "end repeat")
    is_end <- str_starts(t, "end_group") || str_starts(t, "end group") || is_repeat_end
    
    if (is_group || is_repeat) {
      lbl <- r[[label_col_name]] %||% r$name %||% ""
      if (is_group) {
        prev_type <- if (i > 1) tolower(as.character(survey$type[i - 1] %||% "")) else ""
        prev_is_group <- str_starts(prev_type, "begin_group") || str_starts(prev_type, "begin repeat")
        if ( !prev_is_group) { sec_id <- sec_id + 1L; sub_id <- 0L; q_id <- 0L; doc <- add_band(doc, glue("Section {sec_id} : {lbl}"), txt_fp = fp_sec_title) } else { sub_id <- sub_id + 1L; q_id <- 0L; doc <- add_band(doc, glue("Sous-section {sec_id}.{sub_id} : {lbl}"), txt_fp = fp_sub_title) }
      } else if (is_repeat) { doc <- body_add_fpar(doc, fpar(ftext(glue("🔁 Bloc : {lbl}"), fp_block), fp_p = p_default)) }
      rel <- r$relevant %||% NA_character_
      if (!is.na(rel) && nzchar(rel)) { 
        tr <- translate_relevant(rel, lab_map, choices, survey); 
        doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) 
      }
      doc <- add_hrule(doc) ; next
    }
    if (is_end) {
      doc <- body_add_fpar(doc, fpar(ftext("--- Fin du bloc ---", fp_hint), fp_p = p_default))
      doc <- add_hrule(doc)
      next 
    }
    
    if(t != "note") q_id <- q_id + 1L
    
    current_number <- NULL
    if(t != "note") {
      current_number <- glue("{q_id}")
      if (sub_id > 0) { current_number <- glue("{sec_id}.{sub_id}.{q_id}") } else if (sec_id > 0) { current_number <- glue("{sec_id}.{q_id}") }
    }
    
    # Appel à render_question
    doc <- render_question(doc, r, current_number, label_col_name, hint_col_name, choices_map, full_choices_sheet = choices, lab_map = lab_map, full_survey_sheet = survey)
  }
  
  # Écriture du document final vers le chemin temporaire de Shiny
  print(doc, target = output_path)
  return(output_path)
}

# ---------------------------------------------------------------------
# Partie 2 : Fonction d'analyse et de validation d'erreurs XLSForm (COMPLÈTE)
# ---------------------------------------------------------------------
validate_xlsform <- function(filepath) {
  errors <- list()
  
  add_error <- function(sheet, line, type, description, suggestion = NULL) {
    criticite <- case_when(
      type %in% c("Erreur Critique", "Structure", "Nom Invalide", "Référence Invalide", "Duplication", "Données Manquantes") ~ "Critique",
      type %in% c("Amélioration", "Info") ~ "Amélioration",
      TRUE ~ "Autre"
    )
    errors <<- append(errors, list(data.frame(
      Feuille = sheet, Ligne = as.character(line), Type = type, Description = description, 
      Suggestion = suggestion, Criticite = criticite, stringsAsFactors = FALSE
    )))
  }
  
  tryCatch({
    survey <- read_excel(filepath, sheet = "survey")
    choices <- tryCatch(read_excel(filepath, sheet = "choices"), error = function(e) NULL)
    
    names(survey) <- tolower(names(survey))
    label_col_name_survey <- detect_label_col(survey)
    
    if (!all(c("type", "name") %in% names(survey))) {
      missing_cols <- setdiff(c("type", "name"), names(survey))
      add_error("survey", "N/A", "Structure", "Colonnes obligatoires manquantes dans l'onglet 'survey'", glue("Veuillez ajouter les colonnes : {paste(missing_cols, collapse=', ')}."))
    }
    if (is.na(label_col_name_survey)) {
      add_error("survey", "En-tête", "Structure", "Aucune colonne 'label' ou 'label::langue' trouvée.", "Ajoutez au moins une colonne de label pour les questions.")
    }
    
    if ("name" %in% names(survey)) {
      for (i in seq_len(nrow(survey))) {
        q_name <- survey$name[i]
        if (is.na(q_name) || !nzchar(as.character(q_name))) {
          add_error("survey", i + 1, "Données Manquantes", "Le nom de la variable est vide.", "Toutes les questions doivent avoir un nom unique et valide.")
        } else if (!grepl("^[a-zA-Z_][a-zA-Z0-9_]*$", as.character(q_name))) {
          add_error("survey", i + 1, "Nom Invalide", glue("Le nom '{q_name}' contient des caractères non autorisés ou commence par un chiffre."), "Utilisez uniquement des lettres, chiffres, et tirets bas (_).")
        }
      }
      dups <- survey$name[duplicated(survey$name) & !is.na(survey$name)]
      if (length(dups) > 0) {
        for (dup_name in unique(dups)) {
          lines <- which(survey$name == dup_name) + 1
          add_error("survey", paste(lines, collapse=", "), "Duplication", glue("Le nom de variable '{dup_name}' est utilisé plusieurs fois."), "Les noms de variables doivent être uniques.")
        }
      }
    }
    
    if (!is.null(choices)) names(choices) <- tolower(names(choices))
    select_types <- survey %>% filter(str_detect(type, "^select_one |^select_multiple "))
    if (nrow(select_types) > 0) {
      if (is.null(choices) || nrow(choices) == 0) {
        add_error("choices", "N/A", "Données Manquantes", "L'onglet 'choices' est manquant ou vide.", "Ajoutez l'onglet 'choices' avec les listes d'options définies.")
      } else {
        if (!"list_name" %in% names(choices)) {
          add_error("choices", "N/A", "Structure", "La colonne 'list_name' est manquante dans l'onglet 'choices'.", "Cette colonne est essentielle pour lier les choix aux questions.")
        } else {
          required_lists <- str_trim(gsub("select_(one|multiple)\\s+", "", select_types$type))
          available_lists <- unique(choices$list_name)
          for (list_name in required_lists) {
            if (!list_name %in% available_lists) {
              line_num_rows <- which(str_detect(survey$type, list_name))
              line_nums <- paste(line_num_rows + 1, collapse=", ")
              add_error("survey", line_nums, "Référence Invalide", glue("La liste de choix '{list_name}' n'a pas été trouvée dans l'onglet 'choices'."), "Vérifiez l'orthographe ou définissez la liste.")
            }
          }
        }
      }
    }
    
    if (!"hint" %in% names(survey)) add_error("survey", "En-tête", "Amélioration", "La colonne 'hint' est manquante.", "Ajoutez une colonne 'hint' pour les instructions.")
    if (!"relevant" %in% names(survey)) add_error("survey", "En-tête", "Amélioration", "La colonne 'relevant' est manquante.", "Ajoutez une colonne 'relevant' pour la logique de saut.")
    
  }, error = function(e) {
    add_error("Général", "N/A", "Erreur Critique", glue("Impossible de lire le fichier Excel ou onglet manquant : {e$message}"))
  })
  
  if (length(errors) == 0) {
    return(data.frame(Feuille = "N/A", Ligne = "N/A", Type = "Info", Description = "Aucune erreur critique détectée.", Suggestion = "Vous pouvez procéder à la conversion ou au déploiement.", Criticite = "Amélioration", stringsAsFactors = FALSE))
  } else {
    return(bind_rows(errors))
  }
}
# app.R - Partie 3/3 (UI et Logique Serveur Shiny)

# ---------------------------------------------------------------------
# Partie 3 : Interface Utilisateur (UI) et Serveur Shiny 
# ---------------------------------------------------------------------

library(shinythemes)
ui <- fluidPage(theme = shinytheme('flatly'), tags$head(tags$style(HTML('.card {border:1px solid #ddd; border-radius:8px; padding:15px; margin-bottom:15px;} .btn {border-radius:4px;}'))),
  titlePanel("XLSForm Outil de Validation et de Conversion Word"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("file1", "1. Choisir le fichier XLSForm (.xlsx)",
                accept = c(".xlsx")),
      tags$hr(),
      h4("2. Configuration de la langue"),
      actionButton("detect_lang_btn", "Détecter les langues", icon("globe-americas"), class = "btn-secondary"),
      uiOutput("language_select_ui"), 
      tags$hr(),
      h4("3. Actions"),
      actionButton("validate_btn", "Analyser les erreurs", icon("check-circle"), class = "btn-primary"),
      downloadButton("download_word", "Générer Word (.docx)", class = "btn-success"),
      downloadButton("download_excel_report", "Rapport Excel (.xlsx)", class = "btn-info"),
      tags$hr(),
      p("Documentation ODK :"),
      a("[docs.getodk.org](docs.getodk.org)", href="docs.getodk.org", target="_blank")
    ),
    
    mainPanel(
      tabsetPanel(id = "errorTabs",
                  tabPanel("Résumé Général", value = "summary", 
                           h4("Résultats de l'Audit (Résumé)"),
                           p("Le tableau ci-dessous montre le nombre d'erreurs détectées pour chaque catégorie."),
                           DTOutput("summary_error_table"),
                           tags$hr(),
                           h4("Structure du Formulaire"),
                           DTOutput("section_summary_table")
                  ),
                  tabPanel("Critique (Bloquant)", value = "critical", DTOutput("critical_errors")),
                  tabPanel("Structure/Données", value = "data_errors", DTOutput("data_errors_table")),
                  tabPanel("Amélioration/Info", value = "improvement", DTOutput("improvement_table"))
      )
    )
  )
)

server <- function(input, output, session) {
  
  rv <- reactiveValues(errors = NULL, languages = NULL, selected_lang = NULL, survey_data = NULL)
  
  detect_languages <- function(filepath) {
    survey <- read_excel(filepath, sheet = "survey")
    cols <- names(survey)
    lang_cols <- cols[grepl("^label::", tolower(cols))]
    languages <- unique(str_replace(tolower(lang_cols), "^label::", ""))
    if (length(languages) > 0) {
      return(c("Auto", languages))
    } else {
      return(c("Auto (Une seule langue détectée)"))
    }
  }
  
  observeEvent(input$detect_lang_btn, {
    req(input$file1)
    filepath <- input$file1$datapath
    rv$errors <- NULL 
    tryCatch({
      rv$languages <- detect_languages(filepath)
      showNotification("Langues détectées avec succès.", type = 'message', duration = 3) 
    }, error = function(e) {
      rv$languages <- c("Auto (Erreur lecture langues)")
      showNotification(glue("Erreur lors de la détection des langues : {e$message}"), type = "error", duration = NULL)
    })
  })
  
  output$language_select_ui <- renderUI({
    if (!is.null(rv$languages)) {
      selectInput("selected_language", "Choisir la langue d'affichage :", choices = rv$languages, selected = "Auto")
    } else {
      p("Cliquez sur 'Détecter les langues' d'abord.")
    }
  })
  
  observeEvent(input$validate_btn, {
    req(input$file1, input$selected_language)
    filepath <- input$file1$datapath
    
    rv$survey_data <- read_excel(filepath, sheet = "survey") %>% mutate(across(everything(), as.character))
    
    errors_df <- validate_xlsform(filepath)
    rv$errors <- errors_df
    
    updateTabsetPanel(session, "errorTabs", selected = "summary")
  })
  
  render_error_table <- function(data) {
    datatable(data, options = list(pageLength = 10, autoWidth = TRUE, scrollX = TRUE))
  }
  
  # --- Résumé Global (Comme l'image fournie) ---
  output$summary_error_table <- renderDT({
    req(rv$errors)
    
    # Prépare les données pour correspondre exactement aux catégories de l'image
    summary_data_raw <- rv$errors %>%
      mutate(
        Category = case_when(
          str_detect(Description, "choix.*introuvable") ~ "Listes Manquantes",
          str_detect(Type, "Duplication") ~ "Choices Dupliqués", # Utilisé pour les duplications générales ici
          str_detect(Description, "Nom Invalide") ~ "jr.choice-name Invalide", # Mappé au type de l'image
          str_detect(Type, "Référence Invalide") ~ "Références Invalides (relevant)",
          str_detect(Type, "Amélioration") ~ "Problèmes Required", # Mappé au type de l'image (e.g. required_message)
          TRUE ~ "Autres Erreurs"
        )
      ) %>%
      group_by(Category) %>%
      summarise(NombreErreurs = n(), .groups = 'drop') %>%
      rename(TypeErreur = Category)
    
    # Assure que toutes les lignes de l'image sont présentes, même avec 0 erreurs
    all_categories <- c("Listes Manquantes", "Choices Dupliqués", "jr.choice-name Invalide", "Références Invalides (relevant)", "Problèmes Required")
    summary_final <- data.frame(TypeErreur = all_categories, NombreErreurs = 0, stringsAsFactors = FALSE) %>%
      summary_final <- summary_final %>% left_join(summary_data_raw, by = 'TypeErreur') %>% mutate(NombreErreurs = coalesce(NombreErreurs.y, NombreErreurs.x)) %>% select(TypeErreur, NombreErreurs)
    
    datatable(summary_final, options = list(dom = 't', paging = FALSE, ordering = FALSE), 
              colnames = c('TypeErreur', 'NombreErreurs'))
  })
  
  # --- Résumé des Sections et Questions ---
  output$section_summary_table <- renderDT({
    req(rv$survey_data)
    
    survey_data <- rv$survey_data %>% 
      mutate(type_clean = tolower(gsub("\\s", "", type %||% "")))
    
    questions_count <- survey_data %>% 
      filter(!type_clean %in% c(tolower(EXCLUDE_TYPES), "begingroup", "endgroup", "beginrepeat", "endrepeat")) %>% 
      nrow()
    
    sections <- survey_data %>% 
      filter(str_detect(type_clean, "^begingroup|^beginrepeat")) %>%
      mutate(label_col = detect_label_col(survey_data)) %>% mutate(Label = ifelse(!is.na(label_col) & label_col %in% names(survey_data), survey_data[[label_col]], name)) %>% select(name, type, Label) %>%
      mutate(Nb_Questions = 0)
    
    if(nrow(sections) == 0 && questions_count > 0){
      # Utilise htmltools::br() pour simuler un tableau propre pour une ligne unique si pas de sections
      summary_html <- HTML(paste0("<b>Total Questions:</b> ", questions_count, "<br><b>Form Title:</b> ", rv$survey_data$label[1] %||% "N/A"))
      return(renderText(summary_html))
    } else if (nrow(sections) > 0) {
      summary_data <- sections %>% select(Type=type, Nom=name, Label=label)
      return(datatable(summary_data, options = list(pageLength = 10, autoWidth = TRUE, scrollX = TRUE)))
    } else {
      summary_html <- HTML("<i>Aucune question ou section détectée dans le formulaire.</i>")
      return(renderText(summary_html))
    }
  })
  
  
  output$critical_errors <- renderDT({
    req(rv$errors)
    data <- rv$errors %>% filter(Criticite == "Critique") %>% select(-Criticite)
    render_error_table(data)
  })
  
  output$data_errors_table <- renderDT({
    req(rv$errors)
    data <- rv$errors %>% filter(Criticite %in% c("Structure", "Données")) %>% select(-Criticite)
    render_error_table(data)
  })
  
  output$improvement_table <- renderDT({
    req(rv$errors)
    data <- rv$errors %>% filter(Criticite == "Amélioration") %>% select(-Criticite)
    render_error_table(data)
  })
  
  output$download_excel_report <- downloadHandler(
    filename = function() {
      paste("Rapport_erreurs_", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(rv$errors)
      write_xlsx(rv$errors, path = file)
    }
  )
  
  output$download_word <- downloadHandler(
    filename = function() {
      req(input$file1, input$selected_language)
      base_filename <- tools::file_path_sans_ext(input$file1$name)
      lang_suffix <- ifelse(input$selected_language == "Auto", "", paste0("_", input$selected_language))
      paste0(base_filename, "_rendu", lang_suffix, ".docx")
    },
    content = function(file) {
      req(input$file1, input$selected_language)
      showNotification("Génération du document Word en cours...", type = 'message', duration = 5)
      
      tryCatch({
        xlsform_to_wordRev_shiny(
          xlsx = input$file1$datapath, 
          output_path = file,
          selected_language = input$selected_language
        )
        showNotification("Document Word généré avec succès !", type = "success")
      }, error = function(e) {
        showNotification(glue("Erreur lors de la génération Word : {e$message}"), type = "error", duration = NULL)
        file.create(file) 
      })
    }
  )
}

# Lancement de l'application Shiny
shinyApp(ui = ui, server = server)

