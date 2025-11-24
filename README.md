# Outil d'Analyse de Conformité XLSForm

Cet outil R permet d'analyser un fichier XLSForm (format Excel) et de générer un rapport détaillé de toutes les erreurs de conformité détectées.

## Utilisation

1.  **Prérequis :** Assurez-vous d'avoir R installé avec les packages suivants :
    ```R
    install.packages(c("readxl", "writexl", "dplyr", "stringr", "here"))
    ```

2.  **Lancer l'analyse :**
    Placez votre fichier XLSForm dans le dossier racine du projet (ou mettez à jour le chemin dans `run_analysis.R`).

    Exécutez le script principal dans votre console R :
    ```R
    source("run_analysis.R")
    ```

## Résultat

Le script générera un fichier nommé `Rapport_Conformite_XLSForm.xlsx` dans le répertoire de travail, listant toutes les erreurs par feuille, ligne et colonne.
