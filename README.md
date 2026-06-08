# Automatisation de la saisie des factures fournisseurs sur l'ERP Herakles

Projet visant à automatiser le traitement des factures fournisseurs reçues par email, leur extraction d'informations, leur suivi dans un fichier Excel et leur intégration dans l'ERP Herakles.
L'automatisation a permis de réduire le temps de traitement d'environ 12 heures à 1 heure par semaine.

## Traitement de base
- Traitement des factures une par une à la main
- Enregistrement de la facture
- Vérification de la présence du numéro de facture et du numéro de commande
- Saisie des informations sur l'ERP pour créer la facture si le bon de réception existe

## Aujourd'hui
- Enregistrement des factures dans un dossier
- Extraction des informations (numéro de facture, numéro de commande, date de facture, date d'échéance, nom du fournisseur) dans un fichier Excel
- Déplacement des factures dans différents dossiers selon le résultat de son traitement (traité, à_vérfier, à_relancer)
- Vérification manuelle de la fiabilité des informations
- Création de la facture grâce aux informations saisies dans le fichier Excel si le bon de réception existe

## Résultats
- Temps de traitement réduit de 12 h à 1 h par semaine
- Réduction des erreurs de saisie
- Centralisation du suivi des factures
- Meilleure traçabilité des traitements
- Réduction des tâches répétitives

## Vue simplifiée du workflow
```text
Emails Outlook
      ↓
Téléchargement des factures
      ↓
Extraction PDF / OCR
      ↓
Analyse des données
      ↓
Fichier Excel de suivi
      ↓
Automatisation Herakles
      ↓
Création des factures
```

## Architecture
Le projet est composé de plusieurs modules :
- Collecte des factures depuis Outlook via Microsoft Graph
- Extraction du texte depuis les PDF et les factures scannées
- Analyse et identification des informations clés
- Mise à jour du fichier Excel de suivi
- Automatisation de l'ERP Herakles via pywinauto
- Génération des fichiers de suivi et de logs

## Prérequis
- Python 3.11+
- Accès Microsoft Graph
- Accès à l'ERP Herakles
- Tesseract OCR installé sur le poste

## Technologies
- Python
- Bash
- Microsoft Graph API
- PDFplumber
- pytesseract
- openpyxl
- pywinauto

## Fonctionnement
- Lecture des mails et enregistrement des factures via une clé secrète et Microsoft Graph
- Lecture des factures soit en tant que PDF soit en tant qu'image (pour les factures scannées) grâce aux bibliothèques PDFplumber et pytesseract
- Reconnaissance des infos nécessaires avec des regex, des matching et l'IA
- Importation de ces infos dans le fichier Excel grâce à la bibliothèque openpyxl
- Pour la création de la facture sur Herakles :
  * Communication avec l'interface graphique grâce à la bibliothèque pywinauto
  * Lancement de l'application via son chemin et connexion avec les identifiants donnés
  * Placement sur le bon module pour faire la recherche de la commande depuis le numéro de commande indiqué sur la facture
  * Une fois la commande trouvée, vérification de la présence du ou des bon(s) de réception (une facture par bon de réception)
  * Si oui, création de la facture et écriture du montant HT dans le fichier Excel + indication du succès de la création
  * Sinon, indication du manque du bon de réception dans le fichier Excel
- Plusieurs fichiers de suivis sont alimentés à chaque traitement pour simplifier la correction d'erreurs

## Cas gérés
- Factures PDF natives
- Factures scannées (OCR)
- Plusieurs bons de réception pour une même commande
- Plusieurs numéros de commande pour une même facture
- Factures sans numéro de commande détectable
- Factures nécessitant une vérification manuelle
- Absence de bon de réception dans l'ERP

## Limites connues
- Certains formats de factures très spécifiques peuvent nécessiter une vérification manuelle
- La qualité de l'OCR dépend de la qualité du scan reçu
- L'automatisation Herakles repose sur l'interface graphique et peut nécessiter des ajustements en cas de mise à jour du logiciel




# ENGLISH VERSION 
## Supplier Invoice Entry Automation in the Herakles ERP
Project aimed at automating the processing of supplier invoices received by email, extracting information from them, tracking them in an Excel file, and integrating them into the Herakles ERP.
The automation reduced processing time from approximately 12 hours to 1 hour per week.

## Initial Process
* Process invoices one by one manually
* Save the invoice
* Check for the presence of the invoice number and purchase order number
* Enter the information into the ERP to create the invoice if the goods receipt exists

## Today
* Save invoices into a folder
* Extract information (invoice number, purchase order number, invoice date, due date, supplier name) into an Excel file
* Move invoices into different folders depending on the processing result (processed, to_verify, follow_up)
* Manually verify the reliability of the information
* Create the invoice using the information entered in the Excel file if the goods receipt exists

## Results
* Processing time reduced from 12 h to 1 h per week
* Reduction of data entry errors
* Centralized invoice tracking
* Better traceability of processing
* Reduction of repetitive tasks

## Simplified Workflow
```text
Outlook Emails
      ↓
Invoice Download
      ↓
PDF / OCR Extraction
      ↓
Data Analysis
      ↓
Excel Tracking File
      ↓
Herakles Automation
      ↓
Invoice Creation
```

## Architecture
The project consists of several modules:
* Invoice collection from Outlook via Microsoft Graph
* Text extraction from PDFs and scanned invoices
* Analysis and identification of key information
* Updating the Excel tracking file
* Herakles ERP automation via pywinauto
* Generation of tracking and log files

## Prerequisites
* Python 3.11+
* Microsoft Graph access
* Access to the Herakles ERP
* Tesseract OCR installed on the workstation

## Technologies
* Python
* Bash
* Microsoft Graph API
* PDFplumber
* pytesseract
* openpyxl
* pywinauto

## How It Works
* Read emails and save invoices using a secret key and Microsoft Graph
* Read invoices either as PDFs or as images (for scanned invoices) using PDFplumber and pytesseract
* Identify the required information using regex, matching techniques, and AI
* Import this information into the Excel file using the openpyxl library
* For invoice creation in Herakles:
  * Communicate with the graphical interface using the pywinauto library
  * Launch the application via its path and log in with the provided credentials
  * Navigate to the appropriate module to search for the purchase order using the purchase order number found on the invoice
  * Once the purchase order is found, verify the presence of one or more goods receipts (one invoice per goods receipt)
  * If present, create the invoice and write the amount excluding tax into the Excel file + indicate successful creation
  * Otherwise, indicate the missing goods receipt in the Excel file
* Several tracking files are updated during each processing run to simplify error correction

## Supported Cases
* Native PDF invoices
* Scanned invoices (OCR)
* Multiple goods receipts for a single purchase order
* Multiple purchase order numbers for a single invoice
* Invoices without a detectable purchase order number
* Invoices requiring manual verification
* Missing goods receipts in the ERP

## Known Limitations
* Some very specific invoice formats may require manual verification
* OCR quality depends on the quality of the received scan
* Herakles automation relies on the graphical user interface and may require adjustments in the event of software updates

