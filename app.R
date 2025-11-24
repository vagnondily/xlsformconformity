# app.R

# Charger les librairies nécessaires
library(shiny)
library(readxl)
library(dplyr)
library(stringr)
library(openxlsx) 

# --- Logique d'Audit (Fonction Server) ---
server <- function(input, output, session) {
  
  # Reactive expression pour lire les données une fois le fichier téléchargé
  xlsform_data <- reactive({
    req(input$fileUpload) 
    
    file_path <- input$fileUpload$datapath
    
    tryCatch({
      survey_data <- read_excel(file_path, sheet = "survey")
      choices_data <- read_excel(file_path, sheet = "choices")
      return(list(survey = survey_data, choices = choices_data))
      
    }, error = function(e) {
      stop(paste("Erreur de lecture du fichier :", e$message))
    })
  })
  
  # Reactive expression pour exécuter l'audit complet et retourner les résultats
  audit_results <- reactive({
    data_list <- xlsform_data()
    survey <- data_list$survey
    choices <- data_list$choices
    
    vars_defined <- survey$name
    
    # 1. Choices en double (DuplicateChoices)
    dup_choices <- choices %>%
      group_by(list_name, name) %>%
      filter(n() > 1) %>%
      ungroup() 
    
    # 2. Listes manquantes (MissingLists)
    lists_used <- survey %>%
      filter(str_detect(type, "select_")) %>%
      mutate(list_name = str_extract(type, "(?<=select_(one|multiple) )\\S+")) %>%
      pull(list_name)
    missing_lists <- setdiff(lists_used, unique(choices$list_name))
    missing_lists_df <- if (length(missing_lists) > 0) {
      data.frame(list_name = missing_lists, stringsAsFactors = FALSE)
    } else {
      data.frame(list_name = character(0), stringsAsFactors = FALSE)
    }
    
    # 3. Références invalides dans 'relevant' (InvalidRefs)
    invalid_refs <- survey %>%
      filter(str_detect(relevant, "\\$\\{")) %>%
      mutate(refs = str_extract_all(relevant, "\\$\\{([^}]+)\\}")) %>%
      rowwise() %>%
      mutate(
        invalid = list(
          setdiff(
            str_replace_all(unlist(refs), fixed("${"), "") %>% str_replace_all(fixed("}"), ""),
            vars_defined
          )
        )
      ) %>%
      filter(length(invalid) > 0) %>%
      ungroup() 
    
    # 4. 'jr:choice-name' invalide (InvalidJRChoiceName) - CORRIGÉ
    invalid_jr <- survey %>%
      filter(str_detect(calculation, "jr:choice-name")) %>%
      # Utilise un groupe de capture () et extrait le groupe 2 (le nom de la variable)
      mutate(
        second_arg_full = str_extract(calculation, "jr:choice-name\\([^,]+,\\s*'( emulation| )*([^']+)'\\)"),
        second_arg = str_replace(second_arg_full, "jr:choice-name\\([^,]+,\\s*'( emulation| )*([^']+)'\\)", "\\2")
      ) %>%
      filter(!is.na(second_arg) & !second_arg %in% vars_defined) %>%
      select(-second_arg_full) %>% # Nettoie la colonne temporaire
      ungroup()
    
    # 5. Problèmes de 'required' (RequiredIssues)
    req_issues <- survey %>%
      filter(!is.na(required) & !tolower(required) %in% c("yes","no","true()","false()")) %>%
      ungroup() 
    
    # Préparer le résumé 
    summary_data <- data.frame(
      TypeErreur = c("Listes Manquantes", "Choices Dupliqués", "jr:choice-name Invalide", "Références Invalides (relevant)", "Problèmes Required"),
      NombreErreurs = c(nrow(missing_lists_df), nrow(dup_choices), nrow(invalid_jr), nrow(invalid_refs), nrow(req_issues))
    )
    
    return(list(
      summary = summary_data,
      dup_choices = dup_choices,
      missing_lists = missing_lists_df, 
      invalid_jr = invalid_jr,
      invalid_refs = invalid_refs,
      req_issues = req_issues
    ))
  })
  
  # --- Téléchargement du rapport Excel ---
  output$downloadReport <- downloadHandler(
    filename = function() {
      "Audit_XLSForm_Report_R.xlsx"
    },
    content = function(file) {
      audit <- audit_results()
      wb <- createWorkbook()
      
      addWorksheet(wb, "Summary")
      writeData(wb, "Summary", audit$summary)
      addWorksheet(wb, "DuplicateChoices")
      writeData(wb, "DuplicateChoices", audit$dup_choices)
      addWorksheet(wb, "MissingLists")
      writeData(wb, "MissingLists", audit$missing_lists)
      addWorksheet(wb, "InvalidJRChoiceName")
      writeData(wb, "InvalidJRChoiceName", audit$invalid_jr)
      addWorksheet(wb, "InvalidRefs")
      writeData(wb, "InvalidRefs", audit$invalid_refs)
      addWorksheet(wb, "RequiredIssues")
      writeData(wb, "RequiredIssues", audit$req_issues)
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # Affichage du résumé dans l'interface utilisateur
  output$summaryTable <- renderTable({
    req(audit_results())
    audit_results()$summary
  })
  
  # Logique pour afficher le bouton de téléchargement de manière conditionnelle
  output$downloadButtonUI <- renderUI({
    if (!is.null(input$fileUpload) && !is.null(xlsform_data())) {
      downloadButton("downloadReport", "Télécharger le Rapport Excel Complet")
    }
  })
}

# --- Interface Utilisateur (UI) ---
ui <- fluidPage(
  titlePanel("Audit Interactif d'XLSForm"),
  
  sidebarLayout(
    sidebarPanel(
      p("Veuillez télécharger votre fichier XLSForm (au format .xlsx) ici."),
      fileInput("fileUpload", "Choisir le fichier XLSForm",
                multiple = FALSE,
                accept = c(".xlsx")),
      
      uiOutput("downloadButtonUI") 
    ),
    
    mainPanel(
      h3("Résultats de l'Audit (Résumé)"),
      p("Le tableau ci-dessous montre le nombre d'erreurs détectées pour chaque catégorie."),
      tableOutput("summaryTable")
    )
  )
)

# Lancer l'application
shinyApp(ui = ui, server = server)
